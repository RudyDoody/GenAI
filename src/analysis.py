"""Analytics + charts shared by the deliverables.

- K-means segmentation of the synthetic buyer base (P1 Marketing)
- Seasonal-index x linear-trend demand forecast (P4 Operations) - identical
  method to the Excel workbook so numbers reconcile
- Charts: 2x2, segments, forecast, ROI ranges, roadmap gantt, approval workflow
Results are written to data/results.json for the document builders.
"""
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

import facts

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
FIG = ROOT / "data" / "figures"
FIG.mkdir(parents=True, exist_ok=True)

# palette (validated reference palette, light mode) + ink
BLUE, ORANGE, AQUA, YELLOW, VIOLET, RED = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#6250d6", "#e34948"
INK, INK2, MUTED, GRID, SURF = "#0b0b0b", "#52514e", "#8a8984", "#e4e3de", "#ffffff"
SEG_COLORS = {"A": BLUE, "B": ORANGE, "C": AQUA}

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8, "axes.titleweight": "bold",
    "axes.titlesize": 12, "axes.titlecolor": INK, "figure.facecolor": SURF, "axes.facecolor": SURF,
    "savefig.dpi": 200, "savefig.bbox": "tight",
})

FEATURES = ["est_annual_volume_k_pcs", "avg_order_qty_per_style", "styles_per_season",
            "sustainability_requirement_score", "price_sensitivity_1to10",
            "expected_sample_turnaround_days", "digital_engagement_score", "target_fob_eur"]


# ------------------------------------------------------------ segmentation
def segment():
    df = pd.read_csv(DATA / "buyers_synthetic.csv")
    X = df[FEATURES].copy()
    X["est_annual_volume_k_pcs"] = np.log(X["est_annual_volume_k_pcs"])
    X["avg_order_qty_per_style"] = np.log(X["avg_order_qty_per_style"])
    Xs = StandardScaler().fit_transform(X)
    inertia = {k: float(KMeans(n_clusters=k, n_init=20, random_state=7).fit(Xs).inertia_) for k in range(2, 7)}
    km = KMeans(n_clusters=3, n_init=20, random_state=7).fit(Xs)
    df["cluster"] = km.labels_
    prof = df.groupby("cluster")[FEATURES].mean()
    # name clusters by centroid: highest volume -> B, highest digital -> C, remaining -> A
    b = prof["est_annual_volume_k_pcs"].idxmax()
    c = prof.drop(index=b)["digital_engagement_score"].idxmax()
    a = [x for x in prof.index if x not in (b, c)][0]
    name = {a: "A", b: "B", c: "C"}
    df["segment"] = df.cluster.map(name)
    df["segment_name"] = df.segment.map(facts.SEGMENT_NAMES)
    key = pd.read_csv(DATA / "_answer_key_latent_segments.csv")
    agree = float((df.merge(key, on="account_id").eval("segment == _latent")).mean())
    df.drop(columns=["cluster"]).to_csv(DATA / "buyers_segmented.csv", index=False)

    summ = []
    for s in "ABC":
        d = df[df.segment == s]
        act = d[d.relationship_status == "Active customer"]
        summ.append(dict(
            seg=s, name=facts.SEGMENT_NAMES[s], accounts=int(len(d)), active=int(len(act)),
            revenue_lakh=float(act.ww_revenue_fy26_inr_lakh.sum()),
            rev_share=float(act.ww_revenue_fy26_inr_lakh.sum() / df.ww_revenue_fy26_inr_lakh.sum()),
            **{f: float(d[f].mean()) for f in FEATURES},
            top_country=d.country.mode()[0], top_channel=d.preferred_channel.mode()[0],
            gots_share=float(d.certifications_required.str.contains("GOTS").mean()),
            opt_in=float(d.email_opt_in.mean()),
        ))
    # chart: profile heat-table style -> grouped dot chart of z-scores
    z = pd.DataFrame(Xs, columns=FEATURES)
    z["segment"] = df.segment
    zm = z.groupby("segment").mean().loc[list("ABC")]
    labels = ["Annual volume (log)", "Order qty / style (log)", "Styles / season", "Sustainability req.",
              "Price sensitivity", "Sample days expected", "Digital engagement", "Target FOB (EUR)"]
    fig, ax = plt.subplots(figsize=(8.2, 4.4))
    y = np.arange(len(labels))
    ax.axvline(0, color=MUTED, lw=1)
    for i, s in enumerate("ABC"):
        ax.scatter(zm.loc[s].values, y + (i - 1) * 0.22, s=60, color=SEG_COLORS[s], zorder=3,
                   edgecolor=SURF, linewidth=1.5, label=f"{s}: {facts.SEGMENT_NAMES[s]}")
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_xlabel("Segment mean, standard deviations from all-buyer average")
    ax.set_title("Three buyer segments from k-means (k=3) on 300 synthetic accounts", loc="left")
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.38), ncol=3, frameon=False, fontsize=8.5)
    ax.grid(axis="y", visible=False)
    fig.savefig(FIG / "segments_profile.png")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(4.6, 3.0))
    ks = list(inertia)
    ax.plot(ks, [inertia[k] for k in ks], color=BLUE, lw=2, marker="o", ms=7, markeredgecolor=SURF)
    ax.annotate("k = 3 chosen (elbow + business readability)", (3, inertia[3]), xytext=(3.3, inertia[3] + 180),
                fontsize=8.5, color=INK2, arrowprops=dict(arrowstyle="-", color=MUTED))
    ax.set_xlabel("Number of clusters (k)")
    ax.set_ylabel("Within-cluster inertia")
    ax.set_title("Elbow check", loc="left")
    fig.savefig(FIG / "segments_elbow.png")
    plt.close(fig)
    return dict(segments=summ, inertia=inertia, agreement_with_latent=agree)


# ---------------------------------------------------------------- forecast
def si_trend_forecast(y, train_n, horizon):
    """Classical decomposition: 2x12 centred MA ratio seasonal index + linear trend on deseasonalised.
    y: 1-D array of actuals (t = 1..len(y)). Model fit on first train_n points.
    Returns fitted+forecast for t = 1..train_n+horizon, and the seasonal index by calendar position."""
    y = np.asarray(y, float)
    yt = y[:train_n]
    n = len(yt)
    cma = np.full(n, np.nan)
    for i in range(6, n - 6):          # 0-based i -> t = i+1 ; needs i-6..i+6
        cma[i] = (yt[i - 6:i + 6].mean() + yt[i - 5:i + 7].mean()) / 2
    ratio = yt / cma
    cal = np.arange(n) % 12            # data start in January -> cal 0 = Jan
    raw = np.array([np.nanmean(ratio[cal == m]) for m in range(12)])
    si = raw * 12 / raw.sum()
    des = yt / si[cal]
    t = np.arange(1, n + 1)
    slope, intercept = np.polyfit(t, des, 1)
    tt = np.arange(1, n + horizon + 1)
    fc = (intercept + slope * tt) * si[(tt - 1) % 12]
    return fc, si, slope, intercept


def forecast():
    o = pd.read_csv(DATA / "orders_monthly.csv")
    piv = o.pivot(index="month", columns="product_line", values="shipped_pieces")[facts.PRODUCTS]
    months = pd.to_datetime(piv.index)
    N = len(piv)              # 45 (Jan 2023 - Sep 2026)
    TRAIN_BT = N - 12         # backtest: fit to Sep 2025, test Oct 2025 - Sep 2026
    res = {"months": [m.strftime("%Y-%m") for m in months], "lines": {}}
    future = pd.date_range(months[-1] + pd.offsets.MonthBegin(1), periods=12, freq="MS")
    res["future_months"] = [m.strftime("%Y-%m") for m in future]
    bt_tot = np.zeros(12)
    fw_tot = np.zeros(12)
    for p in facts.PRODUCTS:
        y = piv[p].to_numpy()
        fc_bt, si_bt, *_ = si_trend_forecast(y, TRAIN_BT, 12)
        fc_fw, si_fw, slope, icpt = si_trend_forecast(y, N, 12)
        bt = fc_bt[TRAIN_BT:]
        act = y[TRAIN_BT:]
        mape = float(np.mean(np.abs(act - bt) / act))
        res["lines"][p] = dict(actual=y.tolist(), backtest=bt.tolist(), forward=fc_fw[N:].tolist(),
                               mape=mape, si=si_fw.tolist(), slope=float(slope), intercept=float(icpt),
                               bias=float((act - bt).sum() / bt.sum()))
        bt_tot += bt
        fw_tot += fc_fw[N:]
    act_tot = piv.sum(axis=1).to_numpy()
    res["total"] = dict(actual=act_tot.tolist(), backtest=bt_tot.tolist(), forward=fw_tot.tolist(),
                        mape=float(np.mean(np.abs(act_tot[TRAIN_BT:] - bt_tot) / act_tot[TRAIN_BT:])),
                        variance=float(act_tot[TRAIN_BT:].sum() - bt_tot.sum()),
                        variance_pct=float((act_tot[TRAIN_BT:].sum() - bt_tot.sum()) / bt_tot.sum()),
                        forward_sum=float(fw_tot.sum()))
    res["train_bt"] = TRAIN_BT

    # chart
    fig, ax = plt.subplots(figsize=(8.6, 3.8))
    ax.plot(months, act_tot / 1e3, color=INK2, lw=2, label="Actual shipments")
    ax.plot(months[TRAIN_BT:], bt_tot / 1e3, color=BLUE, lw=2, ls="--", label="Back-test forecast (fit to Sep 2025)")
    ax.plot(future, fw_tot / 1e3, color=ORANGE, lw=2, label="Forward forecast Oct 2026 - Sep 2027")
    ax.axvspan(months[TRAIN_BT], months[-1], color=BLUE, alpha=0.06, lw=0)
    ax.text(months[TRAIN_BT], ax.get_ylim()[1] * 0.98, f"  back-test window, MAPE {res['total']['mape']:.1%}",
            va="top", fontsize=8.5, color=INK2)
    ax.set_ylabel("Pieces shipped (thousands)")
    ax.set_title("Monthly shipments, all product lines: actual vs forecast (synthetic data)", loc="left")
    ax.legend(frameon=False, fontsize=8.5, loc="lower left", ncol=3, bbox_to_anchor=(0, -0.3))
    fig.savefig(FIG / "forecast_total.png")
    plt.close(fig)
    return res


# -------------------------------------------------------------------- 2x2
def two_by_two():
    fig, (ax, kx) = plt.subplots(1, 2, figsize=(12.5, 6.6), gridspec_kw=dict(width_ratios=[1.35, 1]))
    ax.axhline(5.5, color=MUTED, lw=1)
    ax.axvline(5.5, color=MUTED, lw=1)
    ax.set_xlim(2, 10.2)
    ax.set_ylim(3, 10.2)
    for (x, y, t) in [(2.15, 10.05, "BIG BETS"), (5.65, 10.05, "PRIORITY / QUICK WINS"),
                      (2.15, 5.35, "PARK"), (5.65, 5.35, "FILL-INS")]:
        ax.text(x, y, t, va="top", fontsize=9, color=INK2, fontweight="bold")
    # spread coincident points so every marker and label is visible
    seen = {}
    for (uid, name, fn, v, f, proto, hz) in facts.USE_CASES:
        k = (v, f)
        n = seen.get(k, 0)
        seen[k] = n + 1
        x, y = f + (n * 0.32 - 0.0), v
        if uid in facts.SHOCK_BORN:
            col, mk = VIOLET, "D"
        elif proto:
            col, mk = BLUE, "o"
        else:
            col, mk = MUTED, "o"
        ax.scatter(x, y, s=150 if proto else 100, color=col, marker=mk, zorder=3, edgecolor=SURF, linewidth=1.5)
        ax.text(x, y - 0.3, uid[2:], fontsize=8, color=INK, ha="center", va="top", fontweight="bold" if proto else "normal")
    ax.scatter([], [], color=BLUE, s=80, label="Prototyped in Phase 2 (P1-P5)")
    ax.scatter([], [], color=MUTED, s=60, label="Roadmap or parked")
    ax.scatter([], [], color=VIOLET, s=60, marker="D", label="Added after Shock E")
    ax.legend(loc="lower left", frameon=True, fontsize=8, facecolor=SURF, edgecolor=GRID)
    ax.set_xlabel("Feasibility (data readiness, cost, skills, time-to-value), 1-10")
    ax.set_ylabel("Business value (revenue, margin, compliance, risk), 1-10")
    ax.set_title("WeaveWorks Exports: AI use cases, value x feasibility", loc="left")
    ax.grid(False)
    kx.axis("off")
    kx.set_xlim(0, 1)
    kx.set_ylim(0, len(facts.USE_CASES) + 1)
    kx.text(0, len(facts.USE_CASES) + 0.6, "Key  (V = value, F = feasibility)", fontsize=9.5, fontweight="bold", color=INK)
    for i, (uid, name, fn, v, f, proto, hz) in enumerate(facts.USE_CASES):
        yy = len(facts.USE_CASES) - i - 0.2
        col = VIOLET if uid in facts.SHOCK_BORN else (BLUE if proto else INK2)
        nm = name.split(" (")[0] if len(name) > 52 else name
        tag = f"  [{proto}]" if proto else ""
        kx.text(0, yy, f"{uid[2:]}", fontsize=8.3, color=col, fontweight="bold")
        kx.text(0.07, yy, f"{nm[:56]}{tag}", fontsize=8.3, color=INK if proto else INK2)
        kx.text(0.98, yy, f"V{v} F{f}", fontsize=8, color=INK2, ha="right")
    fig.tight_layout()
    fig.savefig(FIG / "value_feasibility_2x2.png")
    plt.close(fig)


# ---------------------------------------------------------------- ROI chart
def roi_chart():
    rows = [r for r in facts.roi_rows() if r["included"]]
    rows.sort(key=lambda r: r["roi"][1])
    fig, ax = plt.subplots(figsize=(8.4, 4.6))
    y = np.arange(len(rows))
    for i, r in enumerate(rows):
        ax.plot([r["roi"][0] * 100, r["roi"][2] * 100], [i, i], color=GRID, lw=6, solid_capstyle="round", zorder=1)
        ax.scatter(r["roi"][0] * 100, i, color=RED, s=40, zorder=3, edgecolor=SURF)
        ax.scatter(r["roi"][2] * 100, i, color=AQUA, s=40, zorder=3, edgecolor=SURF)
        ax.scatter(r["roi"][1] * 100, i, color=BLUE, s=90, zorder=4, edgecolor=SURF, linewidth=1.5)
        ax.text(r["roi"][2] * 100 + 25, i, f"base {r['roi'][1]:.0%}", va="center", fontsize=8, color=INK2)
    ax.axvline(0, color=INK2, lw=1)
    ax.set_yticks(y, [f"{r['id']} {r['name']}" for r in rows], fontsize=8.5)
    ax.set_xlabel("18-month ROI on direct use-case cost, %  (red = pessimistic, blue = base, green = optimistic)")
    ax.set_title("ROI range by funded use case (assumed drivers, see ROI_Model.xlsx)", loc="left")
    ax.grid(axis="y", visible=False)
    fig.savefig(FIG / "roi_ranges.png")
    plt.close(fig)


# ------------------------------------------------------------ roadmap gantt
ROADMAP = [
    # label, start, end, horizon, workstream
    ("AI acceptable-use policy, risk register, approval board live", 1, 2, "H1", "Governance"),
    ("AI literacy for all staff; AI champions (1 per dept)", 1, 6, "H1", "People"),
    ("UC04 Buyer support Agent (pilot -> all buyers)", 2, 4, "H1", "Customer"),
    ("UC03 Segmentation + personalised campaigns", 2, 4, "H1", "Marketing"),
    ("UC08 RFQ & proposal generator", 3, 5, "H1", "Sales"),
    ("UC06 Demand forecast in Power BI", 2, 5, "H1", "Operations"),
    ("UC15 Costing & quotation assistant (pulled forward)", 2, 5, "H1", "Merch / Finance"),
    ("Stage-gate 1: CFO review of H1 KPIs, release reserve", 6, 6, "Gate", "Governance"),
    ("UC09 Compliance document assistant (DPP / CSDDD)", 5, 8, "H2", "Compliance"),
    ("UC02 Virtual 3D sampling (pulled forward)", 5, 8, "H2", "Design"),
    ("UC17 'WeaveWorks Studio' buyer portal (shock response)", 6, 10, "H2", "Sales / Design"),
    ("UC11 Tamil voice SOP & training assistant", 7, 10, "H2", "People"),
    ("Stage-gate 2: scale or stop each use case", 12, 12, "Gate", "Governance"),
    ("UC13 Bias-checked JDs & screening (HR stream)", 12, 14, "H3", "HR"),
    ("UC12 Export docs & LC checker - business case", 13, 15, "H3", "Logistics"),
    ("UC10 CV defect: collect image data, go/no-go at M18 (no AI budget)", 14, 18, "H3", "Quality"),
    ("Annual AI audit; refresh roadmap for 2028-29", 16, 18, "H3", "Governance"),
]
H_COL = {"H1": BLUE, "H2": AQUA, "H3": VIOLET, "Gate": INK2}


def gantt():
    fig, ax = plt.subplots(figsize=(10, 6.2))
    for i, (lab, s, e, h, ws) in enumerate(ROADMAP):
        if h == "Gate":
            ax.scatter(s, i, marker="D", s=70, color=INK2, zorder=3)
        else:
            ax.barh(i, e - s + 1, left=s - 0.5, height=0.62, color=H_COL[h], edgecolor=SURF, linewidth=2)
    ax.set_yticks(range(len(ROADMAP)), [r[0] for r in ROADMAP], fontsize=8)
    ax.invert_yaxis()
    ax.set_xticks(range(1, 19), [f"M{m}" for m in range(1, 19)], fontsize=7.5)
    ax.set_xlim(0.4, 18.6)
    for x, t in [(3.5, "H1 Prove & protect\nJan-Jun 2027"), (9.5, "H2 Scale & differentiate\nJul-Dec 2027"),
                 (15.5, "H3 Embed & extend\nJan-Jun 2028")]:
        ax.text(x, -1.1, t, ha="center", fontsize=8.5, fontweight="bold", color=INK2)
    for x in (6.5, 12.5):
        ax.axvline(x, color=MUTED, lw=1, ls=":")
    ax.grid(axis="y", visible=False)
    ax.set_title("18-month AI roadmap (post-shock version)", loc="left", pad=34)
    fig.savefig(FIG / "roadmap_gantt.png")
    plt.close(fig)


# ---------------------------------------------------------- approval flow
def approval_flow():
    fig, ax = plt.subplots(figsize=(10, 3.6))
    ax.axis("off")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 3.6)
    steps = [("1. Request", "Use-case owner\nfills 1-page\nAI request form"),
             ("2. Triage", "AI & Risk Officer\nscores risk tier\n(data, people, money)"),
             ("3. Review", "Tier 1: dept head\nTier 2: + IT/security\nTier 3: + AI Board"),
             ("4. Pilot", "Sandbox, synthetic\ndata, human-in-loop,\nKPI baseline"),
             ("5. Go-live", "Sign-off logged in\nrisk register;\nmonitoring set"),
             ("6. Review", "Quarterly re-test;\nincidents -> stop\nor re-approve")]
    w = 1.45
    for i, (h, b) in enumerate(steps):
        x = 0.15 + i * 1.65
        ax.add_patch(FancyBboxPatch((x, 0.7), w, 2.2, boxstyle="round,pad=0.02,rounding_size=0.12",
                                    fc="#eef4fc" if i % 2 == 0 else "#eaf7f1", ec=GRID))
        ax.text(x + w / 2, 2.55, h, ha="center", fontsize=10, fontweight="bold", color=INK)
        ax.text(x + w / 2, 1.75, b, ha="center", va="center", fontsize=8.2, color=INK2)
        if i < len(steps) - 1:
            ax.annotate("", (x + w + 0.18, 1.8), (x + w + 0.02, 1.8), arrowprops=dict(arrowstyle="-|>", color=INK2))
    ax.text(0.15, 0.3, "Tier 3 (high risk) = personal data of workers/candidates, automated decisions about people, "
            "customer-facing commitments, or anything sent to EU buyers as compliance evidence.",
            fontsize=8.2, color=INK2)
    fig.savefig(FIG / "approval_workflow.png")
    plt.close(fig)


if __name__ == "__main__":
    seg = segment()
    fc = forecast()
    two_by_two()
    roi_chart()
    gantt()
    approval_flow()
    out = dict(segmentation=seg, forecast=fc, roi=facts.roi_rows(), budget=facts.budget_summary(),
               roadmap=ROADMAP)
    (DATA / "results.json").write_text(json.dumps(out, indent=1, default=float))
    for s in seg["segments"]:
        print(s["seg"], s["name"], s["accounts"], s["active"], round(s["revenue_lakh"]), f"{s['rev_share']:.0%}")
    print("agreement", seg["agreement_with_latent"])
    print("MAPE total", fc["total"]["mape"], "var%", fc["total"]["variance_pct"], "fwd", fc["total"]["forward_sum"])
    for p, d in fc["lines"].items():
        print(p, round(d["mape"], 3))
