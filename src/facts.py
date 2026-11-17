"""
Single source of truth for every number used across the Group 4 capstone.

RULE (from the brief): every statistic is either CITED or labelled "assumed".
- Anything about WeaveWorks Exports is ASSUMED (the company is fictional).
- Anything about real organisations/regulations carries a citation in SOURCES.
Change a number here and re-run `python src/build_all.py` to regenerate everything.
"""

GROUP = "Group 4"
CLIENT = "WeaveWorks Exports"
SHOCK_CARD = "E"
SHOCK = ("Competitor move: the main rival launches AI personalisation "
         "and cuts prices 10%.")
RIVAL = "KnitSphere Apparels"  # fictional rival, Tiruppur
TODAY = "17 November 2026"

# ---------------------------------------------------------------- team ----
# Placeholder names - replace with your real names in ONE place (here).
TEAM = [
    {"id": "M1", "name": "Member 1", "role": "Chief Marketing Officer (CMO) + Sales lead",
     "owns": "P1 Marketing pack, P5 B2B proposal deck", "sessions": "4-5"},
    {"id": "M2", "name": "Member 2", "role": "Chief Customer Officer (CCO) + Chief AI & Risk Officer",
     "owns": "P2 Support Agent & testing, Governance pack, ADKAR change plan", "sessions": "6, 11"},
    {"id": "M3", "name": "Member 3", "role": "Chief Financial Officer (CFO)",
     "owns": "P3 Peer annual-report summary, 3-scenario budget, ROI model", "sessions": "7, 10"},
    {"id": "M4", "name": "Member 4", "role": "Chief Operating Officer (COO)",
     "owns": "P4 Demand forecast + DAX variance, supply-disruption decision matrix, roadmap integration", "sessions": "8-9"},
]
CEO_NOTE = ("4-person team: the CEO role (18-month roadmap and board pitch) is shared. "
            "M4 integrates the roadmap; M1 anchors the pitch narrative; every member presents their own function.")

# ------------------------------------------------------- client profile ----
# ALL ASSUMED (fictional company built from the brief: Textile exporter, Tiruppur,
# Rs 120 crore revenue, EU buyers demand compliance, workforce low digital skills).
CO = {
    "revenue_cr": 120.0,              # brief
    "ebitda_margin": 0.09,            # assumed
    "pat_margin": 0.035,              # assumed
    "employees": 1150,                # assumed
    "production_staff": 940,          # assumed
    "merchandisers": 28,              # assumed
    "designers": 9,                   # assumed
    "sales_cs_staff": 14,             # assumed
    "qa_compliance_staff": 32,        # assumed
    "pieces_per_year_mn": 3.8,        # assumed (latest 12 months shipped)
    "capacity_pcs_per_month": 360000, # assumed in-house stitching capacity (4.32 mn/yr)
    "avg_fob_inr": 316,               # assumed (= 120 cr / 3.8 mn pcs)
    "eur_inr": 100.0,                 # assumed planning rate
    "export_share": 0.88,             # assumed
    "eu_share_of_exports": 0.78,      # assumed (rest UK + others)
    "active_buyers": 38,              # assumed
    "sampling_days_now": 21,          # assumed
    "sampling_days_target": 10,       # assumed target
    "buyer_queries_per_month": 1600,  # assumed
    "repetitive_query_share": 0.42,   # assumed
    "rfqs_per_year": 150,             # assumed
    "rfq_win_rate": 0.18,             # assumed
    "rework_rate": 0.045,             # assumed
    "annual_samples": 3600,           # assumed
    "attrition_hires_per_year": 300,  # assumed (production)
    "forecast_mape_now": 0.24,        # assumed (manual spreadsheet forecast)
}
PRODUCTS = ["Basic & fashion T-shirts", "Polos", "Kidswear", "Loungewear"]
SEGMENT_NAMES = {
    "A": "Sustainable mid-market brands",
    "B": "Volume value retailers",
    "C": "Agile DTC & emerging labels",
}
SEGMENT_REV_SHARE = {"A": 0.48, "B": 0.37, "C": 0.15}  # assumed, of WeaveWorks revenue

# ------------------------------------------------------------- budget -----
BUDGET_LAKH = 50.0
HORIZON_MONTHS = 18
ROADMAP_START = "January 2027"   # M1 = Jan 2027, M18 = Jun 2028

# Cross-cutting ("hidden") costs, Rs lakh over 18 months - assumed.
HIDDEN_COSTS = [
    ("Training & change management (ADKAR, AI champions, Tamil-language sessions)", 5.0),
    ("Human oversight (part-time AI governance lead + output reviewers)", 4.0),
    ("Security, DPDP/GDPR legal review, vendor due diligence, annual AI audit", 2.0),
]
CONTINGENCY_LAKH = 4.5           # ~9% of programme; plus unallocated reserve released at M9 stage-gate

# ----------------------------------------------------- use-case portfolio --
# Value / feasibility scored 1-10 by the team (assumed, see 2x2 rationale).
# Prototype flag: built in Phase 2.
USE_CASES = [
    # id, name, function, value, feasibility, prototype, horizon
    ("UC01", "Gen-AI design variants & tech-pack drafting", "Design", 7, 6, None, "H2"),
    ("UC02", "Virtual 3D sampling with AI renders", "Design / Sales", 9, 5, None, "H2"),
    ("UC03", "Buyer segmentation & personalised outreach", "Marketing", 7, 8, "P1", "H1"),
    ("UC04", "Buyer support Agent (orders, docs, policy FAQs)", "Customer service", 8, 8, "P2", "H1"),
    ("UC05", "Finance copilot: peer benchmarking & budget scenarios", "Finance", 5, 9, "P3", "H1"),
    ("UC06", "Demand forecasting & capacity planning", "Operations", 8, 7, "P4", "H1"),
    ("UC07", "Supply-disruption decision support (yarn, freight)", "Operations / Sourcing", 7, 7, "P4", "H1"),
    ("UC08", "RFQ & B2B proposal generator", "Sales", 8, 8, "P5", "H1"),
    ("UC09", "Compliance document assistant (DPP, CSDDD, audits)", "Compliance", 9, 6, None, "H2"),
    ("UC10", "Computer-vision fabric & garment defect detection", "Quality", 8, 4, None, "Deferred"),
    ("UC11", "Tamil voice SOP & training assistant for floor staff", "HR / Production", 7, 6, None, "H2"),
    ("UC12", "Export documentation & LC discrepancy checker", "Logistics / Finance", 6, 7, None, "H3"),
    ("UC13", "Bias-checked JDs & resume screening", "HR", 4, 7, None, "H3"),
    ("UC14", "Predictive maintenance for knitting/sewing machines", "Maintenance", 6, 3, None, None),
    ("UC15", "AI costing & quotation assistant (yarn + CMT)", "Merchandising / Finance", 8, 7, None, "H1"),
    ("UC16", "Receivables & FX exposure alerts", "Finance", 5, 6, None, None),
    # Added after Shock E (rival AI personalisation + 10% price cut):
    ("UC17", "Personalised buyer portal 'WeaveWorks Studio' (AI variants + instant quote + traceability)", "Sales / Design", 9, 6, None, "H2"),
]
SHOCK_BORN = {"UC17"}

# ----------------------------------------------------------- ROI model ----
# For each funded use case. Money in Rs lakh. Annual driver volumes are
# ASSUMED; pessimistic / base / optimistic. Benefit over the 18-month window
#   = driver x unit_value x months_live / 12
# months_live = 18 - golive_month + 1
# Cost = one_time + run_per_month x months_live
ROI = [
    # id, short name, key driver, unit, unit_value_lakh, (pess, base, opt), golive_month, one_time, run_pm, included, pre_shock_golive
    ("UC04", "Buyer support Agent", "Merchandiser hours saved / yr", "hrs", 0.0045, (900, 1700, 2400), 3, 1.8, 0.10, 1, 3),
    ("UC03", "Segmentation & personalised outreach", "New buyer accounts won / yr", "accounts", 7.5, (0.5, 1.5, 3), 3, 1.4, 0.08, 1, 3),
    ("UC08", "RFQ & proposal generator", "Win-rate uplift (percentage points)", "pp", 3.6, (0.5, 2, 4), 4, 1.2, 0.06, 1, 4),
    ("UC06", "Demand forecasting (Power BI)", "Air-freight shipments avoided / yr", "shipments", 1.6, (2, 5, 9), 4, 2.6, 0.08, 1, 4),
    ("UC15", "Costing & quotation assistant", "Margin recovered (basis points of revenue)", "bps", 1.2, (5, 12, 25), 4, 2.4, 0.07, 1, 8),
    ("UC09", "Compliance document assistant", "Compliance hours saved / yr", "hrs", 0.006, (500, 900, 1400), 7, 2.4, 0.10, 1, 7),
    ("UC02", "Virtual 3D sampling", "Physical samples avoided / yr", "samples", 0.025, (250, 600, 1000), 7, 3.6, 0.20, 1, 10),
    ("UC11", "Tamil voice SOP assistant", "Ramp-up days saved per new hire x hires", "hire-days", 0.004, (600, 1500, 2400), 9, 1.8, 0.08, 1, 9),
    ("UC10", "CV defect detection (1-line pilot)", "Rejected pieces avoided / yr", "pieces", 0.0018, (3000, 8000, 14000), 14, 3.4, 0.10, 0, 14),
    ("UC13", "Bias-checked JDs & screening", "Recruiter hours saved / yr", "hrs", 0.004, (150, 300, 500), 13, 0.5, 0.03, 1, 13),
    ("UC17", "Personalised buyer portal (shock response)", "At-risk revenue retained (Rs crore / yr)", "Rs cr", 15.0, (0.5, 1.5, 3.0), 9, 2.6, 0.12, 1, None),
]
SCENARIOS = ["Pessimistic", "Base", "Optimistic"]


def months_live(golive):
    return HORIZON_MONTHS - golive + 1


def _payback(one, run, ben_pm, golive):
    """Programme month in which cumulative benefit first covers cumulative cost (None if not in 18m)."""
    cum = -one
    for m in range(golive, HORIZON_MONTHS + 1):
        cum += ben_pm - run
        if cum >= 0:
            return m
    return None


def roi_rows():
    out = []
    for (uid, name, driver, unit, uv, vols, golive, one, run, inc, pre_gl) in ROI:
        ml = months_live(golive)
        cost = one + run * ml
        ben = [v * uv * ml / 12 for v in vols]
        be_driver = cost / (uv * ml / 12)          # annual driver needed to break even in window
        out.append(dict(id=uid, name=name, driver=driver, unit=unit, unit_value=uv, vols=vols,
                        golive=golive, months_live=ml, one_time=one, run_pm=run, cost=cost,
                        benefit=ben, net=[b - cost for b in ben],
                        roi=[(b - cost) / cost for b in ben], breakeven_driver=be_driver,
                        included=bool(inc), pre_shock_golive=pre_gl,
                        payback_month=[_payback(one, run, v * uv / 12, golive) for v in vols]))
    return out


def budget_summary():
    rows = roi_rows()
    uc_cost = sum(r["cost"] for r in rows if r["included"])
    hidden = sum(v for _, v in HIDDEN_COSTS)
    total = uc_cost + hidden + CONTINGENCY_LAKH
    return dict(uc_cost=uc_cost, hidden=hidden, contingency=CONTINGENCY_LAKH,
                total=total, headroom=BUDGET_LAKH - total)


# ------------------------------------------------------------- sources ----
SOURCES = {
    "SPAL_FY25_RESULTS": "S.P. Apparels Ltd, Q4 & FY25 results filing to BSE and investor presentation (May 2025), s-p-apparels.com / bseindia.com",
    "SPAL_AR25": "S.P. Apparels Ltd, Annual Report 2024-25 (s-p-apparels.com, Aug 2025)",
    "SPAL_BRSR24": "S.P. Apparels Ltd, Business Responsibility & Sustainability Report 2023-24",
    "SPAL_CONCALL": "S.P. Apparels Ltd, Q4 FY25 earnings call (May 2025) and Q1/Q2 FY26 call summaries",
    "DPDP": "Digital Personal Data Protection Act, 2023 (Act No. 22 of 2023), Government of India",
    "DPDP_RULES": "Digital Personal Data Protection Rules, 2025, notified by MeitY in November 2025 (phased commencement)",
    "CERTIN": "CERT-In Directions under Section 70B(6) of the IT Act, 28 April 2022 (6-hour incident reporting)",
    "EU_AI_ACT": "Regulation (EU) 2024/1689 (Artificial Intelligence Act); Art. 50 transparency duties apply from 2 Aug 2026",
    "GDPR": "Regulation (EU) 2016/679 (GDPR), Art. 3(2) territorial scope",
    "CSDDD": "Directive (EU) 2024/1760 (CSDDD), as amended by the 2025 'Omnibus' simplification package",
    "ESPR": "Regulation (EU) 2024/1781 (Ecodesign for Sustainable Products Regulation) - Digital Product Passport; textiles a priority product group",
    "EU_FLR": "Regulation (EU) 2024/3015 prohibiting products made with forced labour (applies from Dec 2027)",
    "REACH": "Regulation (EC) 1907/2006 (REACH), Annex XVII restrictions incl. azo dyes",
    "INDIA_UK_FTA": "India-UK Comprehensive Economic and Trade Agreement, signed July 2025",
    "ADKAR": "Hiatt, J. (2006). ADKAR: A Model for Change in Business, Government and our Community. Prosci.",
}
