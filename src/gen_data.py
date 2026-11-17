"""Generate ALL synthetic datasets (ground rule: no real personal/customer data).

Outputs (in /data and copied into deliverable folders):
  buyers_synthetic.csv   - 300 EU/UK B2B buyer accounts (38 active customers + prospects/lapsed)
  orders_monthly.csv     - monthly shipped pieces by product line, Jan 2023 - Sep 2026
  resumes are not generated (HR stream deferred to roadmap H3)
"""
import numpy as np
import pandas as pd
from pathlib import Path
import facts

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
DATA.mkdir(exist_ok=True)
rng = np.random.default_rng(20261117)

COUNTRIES = {
    "A": (["Germany", "Netherlands", "Denmark", "Sweden", "France", "UK"], [0.30, 0.18, 0.14, 0.13, 0.13, 0.12]),
    "B": (["Germany", "UK", "Spain", "France", "Poland", "Italy"], [0.28, 0.24, 0.16, 0.14, 0.10, 0.08]),
    "C": (["UK", "Netherlands", "Germany", "France", "Ireland", "Belgium"], [0.30, 0.20, 0.18, 0.14, 0.10, 0.08]),
}
BUYER_TYPE = {
    "A": (["Brand", "Brand", "Retailer"], None),
    "B": (["Retailer", "Discount retailer", "Wholesaler"], None),
    "C": (["DTC label", "DTC label", "Brand"], None),
}
# latent segment parameters: mean, sd
P = {
    #        volume k-pcs/yr   order qty/style  styles/season  sust score  price sens  sample days exp  digital  fob EUR
    "A": dict(vol=(260, 90), oq=(2200, 700), sty=(45, 12), sus=(84, 7), ps=(4.5, 1.2), sd=(14, 3), dig=(58, 12), fob=(4.6, 0.6), terms=(60, 10)),
    "B": dict(vol=(900, 300), oq=(9500, 2500), sty=(28, 8), sus=(52, 10), ps=(8.6, 0.8), sd=(21, 4), dig=(35, 10), fob=(2.6, 0.35), terms=(90, 15)),
    "C": dict(vol=(55, 25), oq=(600, 250), sty=(70, 18), sus=(70, 12), ps=(6.0, 1.3), sd=(8, 2), dig=(82, 8), fob=(5.4, 0.8), terms=(30, 8)),
}
CERTS = {"A": ["GOTS", "OEKO-TEX 100", "GRS", "Fair Wear"], "B": ["OEKO-TEX 100", "BSCI/amfori"], "C": ["GOTS", "OEKO-TEX 100"]}


def gen_buyers(n=300):
    seg = rng.choice(["A", "B", "C"], size=n, p=[0.38, 0.30, 0.32])
    rows = []
    for i, s in enumerate(seg, 1):
        p = P[s]
        c_list, c_p = COUNTRIES[s]
        def nrm(k, lo):
            m, sd = p[k]
            return max(lo, rng.normal(m, sd))
        certs = [c for c in CERTS[s] if rng.random() < 0.75] or [CERTS[s][0]]
        rows.append(dict(
            account_id=f"WW-B{i:04d}",
            country=rng.choice(c_list, p=c_p),
            buyer_type=rng.choice(BUYER_TYPE[s][0]),
            est_annual_volume_k_pcs=round(nrm("vol", 10)),
            avg_order_qty_per_style=int(round(nrm("oq", 150), -1)),
            styles_per_season=int(nrm("sty", 6)),
            sustainability_requirement_score=int(min(100, nrm("sus", 10))),
            price_sensitivity_1to10=round(min(10, nrm("ps", 1)), 1),
            expected_sample_turnaround_days=int(nrm("sd", 4)),
            digital_engagement_score=int(min(100, nrm("dig", 5))),
            target_fob_eur=round(nrm("fob", 1.6), 2),
            payment_terms_days=int(round(nrm("terms", 15) / 15) * 15),
            certifications_required=";".join(certs),
            preferred_channel=rng.choice(["Email", "LinkedIn", "Trade fair", "Agent"],
                                         p={"A": [0.35, 0.30, 0.25, 0.10], "B": [0.30, 0.10, 0.30, 0.30], "C": [0.30, 0.50, 0.10, 0.10]}[s]),
            _latent=s,
        ))
    df = pd.DataFrame(rows)
    # relationship status: 38 active customers, ~40 lapsed, rest prospects
    status = np.array(["Prospect"] * n, dtype=object)
    for s, k in {"A": 15, "B": 9, "C": 14}.items():
        idx = df.index[df._latent == s].to_numpy()
        pick = rng.choice(idx, size=k, replace=False)
        status[pick] = "Active customer"
    rest = np.where(status == "Prospect")[0]
    status[rng.choice(rest, size=40, replace=False)] = "Lapsed"
    df["relationship_status"] = status
    df["years_with_weaveworks"] = [int(rng.integers(1, 12)) if st == "Active customer" else
                                   (int(rng.integers(1, 6)) if st == "Lapsed" else 0) for st in status]
    # WeaveWorks revenue per active account scaled to segment shares of Rs 120 cr
    df["ww_revenue_fy26_inr_lakh"] = 0.0
    for s, share in facts.SEGMENT_REV_SHARE.items():
        m = (df._latent == s) & (df.relationship_status == "Active customer")
        w = df.loc[m, "est_annual_volume_k_pcs"].to_numpy() * rng.uniform(0.6, 1.4, m.sum())
        df.loc[m, "ww_revenue_fy26_inr_lakh"] = np.round(w / w.sum() * share * facts.CO["revenue_cr"] * 100, 1)
    df["email_opt_in"] = rng.random(n) < np.where(df.relationship_status == "Prospect", 0.55, 0.9)
    return df


def gen_orders():
    months = pd.date_range("2023-01-01", "2026-09-01", freq="MS")
    # seasonality of SHIPMENTS to EU (assumed): SS shipping Jan-Apr, AW shipping Jul-Sep
    season = np.array([1.10, 1.18, 1.22, 1.08, 0.86, 0.80, 1.05, 1.15, 1.08, 0.88, 0.78, 0.82])
    mix = {"Basic & fashion T-shirts": 0.44, "Polos": 0.18, "Kidswear": 0.24, "Loungewear": 0.14}
    growth = {"Basic & fashion T-shirts": 0.045, "Polos": 0.02, "Kidswear": 0.09, "Loungewear": 0.12}
    base_month = facts.CO["pieces_per_year_mn"] * 1e6 / 12 / 1.10  # so FY26 ~3.8 mn
    rows = []
    for t, d in enumerate(months):
        for p, sh in mix.items():
            trend = (1 + growth[p]) ** (t / 12)
            s = season[d.month - 1]
            if p == "Loungewear":  # stronger AW peak
                s = s * (1.12 if d.month in (8, 9, 10) else 0.97)
            noise = rng.normal(1, 0.06)
            shock = 0.85 if (d.year == 2024 and d.month in (2, 3)) else 1.0  # Red Sea rerouting delays (assumed impact)
            rows.append(dict(month=d.strftime("%Y-%m-%d"), product_line=p,
                             shipped_pieces=int(base_month * sh * trend * s * noise * shock)))
    df = pd.DataFrame(rows)
    # calibrate: latest 12 months (Oct 2025 - Sep 2026) = annual volume implied by Rs 120 cr revenue
    last12 = df[df.month >= "2025-10-01"].shipped_pieces.sum()
    df["shipped_pieces"] = (df.shipped_pieces * facts.CO["pieces_per_year_mn"] * 1e6 / last12).round().astype(int)
    return df


if __name__ == "__main__":
    b = gen_buyers()
    b.drop(columns=["_latent"]).to_csv(DATA / "buyers_synthetic.csv", index=False)
    b[["account_id", "_latent"]].to_csv(DATA / "_answer_key_latent_segments.csv", index=False)
    o = gen_orders()
    o.to_csv(DATA / "orders_monthly.csv", index=False)
    print(b.relationship_status.value_counts().to_dict())
    print(b.groupby("_latent").ww_revenue_fy26_inr_lakh.sum().round(1).to_dict())
    o["year"] = o.month.str[:4]
    print(o.groupby("year").shipped_pieces.sum().to_dict())
