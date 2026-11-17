"""Phase 1 - Diagnose: Opportunity scan (2 pages) + value x feasibility 2x2."""
import facts
from docx_lib import Doc, FIG, ROOT

VALUE_CHAIN = [
    ("1. Win buyers", "Trade fairs (e.g. Munich Fabric Start), agents, LinkedIn; 38 active EU/UK buyers",
     "Founder-led selling; generic catalogues; RFQ replies take 3-5 days", "UC03, UC08, UC17"),
    ("2. Design & sampling", "Tech-packs, proto/fit/SMS samples couriered to Europe",
     "21-day sampling cycle; ~3,600 samples/yr; buyers now expect 10 days", "UC01, UC02"),
    ("3. Costing & order", "Yarn + knitting + dyeing + CMT costing in Excel; price negotiation",
     "Yarn-price swings make quotes stale; margin leakage on small orders", "UC15"),
    ("4. Source materials", "Cotton yarn from Coimbatore/Tiruppur spinners; dyeing on job-work",
     "Single-source yarn risk; ZLD dye-house capacity; freight disruption", "UC07"),
    ("5. Make", "Knitting, cutting, stitching, finishing; ~940 production staff",
     "High attrition; Tamil/Hindi-speaking workforce with low digital skills", "UC11, UC14"),
    ("6. Quality & compliance", "Inline/final QC; GOTS, OEKO-TEX, buyer social audits",
     "4.5% rework; growing EU due-diligence & Digital Product Passport questionnaires", "UC09, UC10"),
    ("7. Ship & document", "Export docs, LC, Chennai/Tuticorin ports; EU transit",
     "Document discrepancies delay payment; Red Sea rerouting adds transit days", "UC12, UC06"),
    ("8. Serve buyers", "Merchandisers answer order-status, certificate and claims queries",
     "~1,600 queries/month, ~42% repetitive; replies across time zones", "UC04"),
    ("Support: Finance & HR", "FX, receivables, budgeting; hiring ~300 operators/yr",
     "Manual reporting; biased/unclear JDs; slow onboarding", "UC05, UC16, UC13"),
]

WHY_PICK = [
    ("P1 Marketing", "UC03", "Top-right. Synthetic CRM data is easy to build; directly feeds the shock response (personalisation)."),
    ("P2 Customer support", "UC04", "Top-right. 42% repetitive queries; low data risk if the Agent only uses a curated knowledge base."),
    ("P3 Finance", "UC05", "Bottom-right fill-in: low value alone but feasible in days, and it gives the CFO a benchmark and the budget envelope for every other case."),
    ("P4 Operations", "UC06 + UC07", "Top-right. Forecast error drives air-freight and yarn over-buying; Ops chosen over HR because buyers' on-time delivery is the No.1 retention driver (assumed from merchandiser interviews, to be validated)."),
    ("P5 Sales", "UC08", "Top-right. Faster, tailored proposals for the largest segment; reuses P1 personas and P2 knowledge base."),
]


def build():
    d = Doc("Opportunity Scan", "Phase 1 - Diagnose: value chain, 17 AI use cases, value x feasibility 2x2",
            "Whole team (CEO role shared); facilitator: Member 4 (COO)", "01-Diagnose")
    d.h1("1. The client in one paragraph")
    co = facts.CO
    d.p(f"**{facts.CLIENT}** is a Tiruppur knitwear exporter with **Rs {co['revenue_cr']:.0f} crore** revenue (brief), "
        f"about **{co['employees']:,} employees** and **{co['pieces_per_year_mn']} million pieces** a year (assumed). "
        f"About {co['export_share']:.0%} of sales are exports, mostly to EU fashion brands and retailers (assumed). "
        "Two tensions shape every recommendation: **EU buyers demand compliance evidence** (due diligence, traceability, "
        "chemical safety) and the **workforce has low digital skills**. The CEO has **Rs 50 lakh and 18 months**.")
    d.h1("2. Value chain map: where AI can help")
    d.table(["Stage", "What happens today", "Pain point (assumed)", "AI use cases"],
            VALUE_CHAIN, widths=[2.6, 5.0, 6.0, 3.4], size=8.5, first_col_bold=True)
    d.h1("3. Scoring method")
    d.bullets([
        "**Value (1-10):** revenue or margin impact, compliance/risk reduction, buyer retention, and fit with the CEO's goals.",
        "**Feasibility (1-10):** data readiness, cost against the Rs 50 lakh envelope, skills needed, time-to-value under 6 months, and vendor maturity.",
        "Each member scored independently; we used the median and discussed any gap of more than 2 points. Scores are team judgements (assumed), not measured data.",
    ])
    d.h1("4. The 2x2 and the five prototypes")
    d.image(FIG / "value_feasibility_2x2.png", caption="Figure 1. Value x feasibility (team scores, assumed). Blue = prototyped in Phase 2.")
    d.table(["Prototype", "Use case", "Why it was picked"], WHY_PICK, widths=[3.0, 2.4, 11.6], size=8.5, first_col_bold=True)
    d.callout(["**Big bets kept for later:** UC02 virtual sampling and UC09 compliance assistant are high-value but need "
               "3D-skills and clean product data first, so they start in H2. UC10 computer-vision QC needs labelled defect "
               "images we do not have yet, so it is deferred.",
               "**Parked:** UC14 predictive maintenance (machines lack sensors) and UC16 FX alerts (bank tools already cover it)."],
              kind="info", title="Not picked, and why")
    d.page_break()
    d.h1("Annex A. Use-case long list (17)")
    rows = [(u[0], u[1], u[2], u[3], u[4], u[5] or "", u[6] or "Parked") for u in facts.USE_CASES]
    d.table(["ID", "Use case", "Function", "V", "F", "Proto", "Horizon"], rows,
            widths=[1.2, 7.4, 3.6, 0.8, 0.8, 1.2, 2.0], size=8)
    d.p("UC17 was added in week 3, after Shock E (rival launches AI personalisation, cuts prices 10%). "
        "UC10 moved from H3 to 'deferred' to fund it. See 03_Stress_Test.", size=9, italic=True)
    d.sources([], extra=["Use-case scores: team workshop, Session 4 (assumed judgements)."])
    return d.save(ROOT / "01_Diagnose" / "Opportunity_Scan.docx")


if __name__ == "__main__":
    print(build())
