"""P3 - Finance: AI-drafted summary of a listed peer's annual report + verification table; 3-scenario budget (xlsx)."""
from openpyxl import Workbook
from openpyxl.styles import Alignment

import facts
from docx_lib import Doc, ROOT
from xlsx_lib import (INPUT_FILL, LAKH, PCT, SUB_FILL, cell, header, note, title, widths, font)

P3 = ROOT / "02_Build" / "P3_Finance"
PEER = "S.P. Apparels Ltd (NSE: SPAL, BSE: 540048)"

# --------------------------------------------------------------- summary
SUMMARY = [
    ("Who they are", "Integrated knitwear maker and exporter headquartered near Tiruppur (Avinashi), Tamil Nadu: knitting, processing, "
     "garmenting; strongest in infant and children's wear, expanding into adult wear. Small domestic retail arm; UK subsidiary; "
     "acquired Young Brand Apparel; building capacity in Sri Lanka."),
    ("Growth", "FY2024-25 consolidated total income about Rs 1,407 crore, up about 27% on FY2023-24 (about Rs 1,104 crore), helped by acquisitions."),
    ("Profitability", "Consolidated EBITDA about Rs 200 crore (about 14% margin). Consolidated PAT about Rs 95 crore (FY24: about Rs 90 crore). "
     "Management's 'adjusted' garment-business view shows a higher margin (about 16-17%), on a different basis."),
    ("Balance sheet", "Debt-to-equity rose to about 0.30 (from about 0.16) as term loans and working capital funded acquisitions and capex."),
    ("Strategy", "Capacity additions (Salem; dyeing capacity; Sri Lanka scaling toward ~2,000 machines by FY27); India-UK FTA (signed July 2025) seen as a tailwind; "
     "retail division small and soft (FY25 about Rs 79 crore vs Rs 83 crore)."),
    ("Risks named", "Cotton and yarn price volatility (stable in FY25 per management); currency (hedged with forwards); tariff uncertainty delaying "
     "India capacity expansion; integration of acquisitions; retail weakness."),
    ("ESG", "About 82% of electricity from renewable sources (FY25 BRSR); exports were about 93% of turnover (FY25 BRSR)."),
]

# claim, AI draft said, verified value, source key, status, page ref in official AR 2024-25 PDF (286 pages)
VERIFY = [
    ("Consolidated total income FY25", "Rs 1,407 crore", "Rs 1,407.3 crore (Rs 14,073.26 mn)", "SPAL_AR25", "Verified", "p.10 (Directors' Report); p.12; p.186 (consol. P&L)"),
    ("Revenue growth FY25 vs FY24", "+27.5%", "+27.5% (Rs 14,073.3 mn vs Rs 11,036.7 mn)", "SPAL_AR25", "Verified", "p.10"),
    ("Revenue from operations FY25", "(not stated)", "Rs 1,395.1 crore (Rs 13,951.34 mn); total income includes other income", "SPAL_AR25", "Added", "p.12; p.186"),
    ("Consolidated PAT FY25", "Rs 83.5 crore", "Rs 95.1 crore consolidated (Rs 951.02 mn); Rs 83.5 crore is STANDALONE", "SPAL_AR25", "CORRECTED - AI mixed bases", "p.10; p.12; p.186"),
    ("EBITDA margin FY25", "16.7%", "14.2% consolidated (EBITDA Rs 2,000.0 mn / Rs 14,073.3 mn); 16.7% is management's adjusted figure", "SPAL_AR25", "CORRECTED - basis clarified", "p.10"),
    ("EPS FY25", "Rs 37.88", "Rs 37.90 basic and diluted (FY24 Rs 35.72)", "SPAL_AR25", "CORRECTED - minor", "p.10; p.187"),
    ("Debt position", "'Virtually debt-free'", "Debt-equity ratio 0.30 (FY24 0.16), up on new term loans and working capital", "SPAL_AR25", "CORRECTED - AI claim false", "p.11 (Key Ratios)"),
    ("Order book", "Rs 442 crore", "Rs 442 crore (Chairman's message)", "SPAL_AR25", "Verified", "p.5"),
    ("Renewable electricity share", "82%", "82% of total power consumption from windmills and rooftop solar", "SPAL_AR25", "Verified", "p.70 (BRSR)"),
    ("Export share of turnover", "'About 95%'", "About 93% of total turnover (FY25 BRSR)", "SPAL_AR25", "CORRECTED - value", "p.38 (BRSR)"),
    ("India-UK FTA", "'In force since 2024'", "Report calls it 'recent'; signed July 2025, in force only after ratification", "INDIA_UK_FTA", "CORRECTED - date", "p.5; p.8"),
    ("Retail division revenue", "Rs 79.4 crore", "Rs 79.4 crore FY25 vs Rs 82.9 crore FY24 (SP Retail Ventures)", "SPAL_AR25", "Verified", "p.5"),
]

BENCH = [
    ("Revenue (Rs crore)", "about 1,407", f"{facts.CO['revenue_cr']:.0f}", "Peer is ~12x our size: compare ratios, not absolutes"),
    ("EBITDA margin", "about 14%", f"{facts.CO['ebitda_margin']:.0%} (assumed)", "Our 5-point gap = Rs 6 crore a year; costing (UC15) and sampling (UC02) attack it"),
    ("PAT margin", "about 6.8%", f"{facts.CO['pat_margin']:.1%} (assumed)", "Thin margin: the AI budget must pay back inside 18 months in the base case"),
    ("Export share", "about 93%", f"{facts.CO['export_share']:.0%} (assumed)", "Both fully exposed to EU/UK buyer rules"),
    ("Renewable electricity", "82%", "about 30% (assumed)", "EU buyers ask for this; a data point the compliance assistant (UC09) must report"),
    ("Capex / expansion", "Salem, dyeing, Sri Lanka", "None planned", "We cannot win on scale; we must win on speed, data and service"),
]

LOG = [
    dict(title="Draft the annual-report summary",
         prompt="Summarise S.P. Apparels Ltd's FY2024-25 annual report for a CFO in 200 words: revenue, growth, EBITDA, PAT, debt, strategy, "
                "risks and ESG. Give figures in Rs crore and state consolidated or standalone for every number.",
         raw="'S.P. Apparels reported revenue of Rs 1,407 crore (+27.5%) and net profit of Rs 83.5 crore. EBITDA margin was a healthy 16.7%. "
             "The company is virtually debt-free and exports about 95% of turnover. The India-UK FTA, in force since 2024, ...'",
         edited="See section 1 summary: consolidated PAT Rs 95.1 crore; EBITDA about 14% consolidated (16.7% is adjusted); D/E about 0.30; "
                "exports about 93%; FTA signed July 2025.",
         why="Despite the instruction, the model mixed standalone PAT with consolidated revenue, took management's adjusted margin as reported, "
             "and asserted 'debt-free' and an FTA date that are false. Six of twelve claims needed correction (see verification table)."),
    dict(title="Build the budget scenarios",
         prompt="For a Rs 50 lakh, 18-month Gen-AI programme at an Indian SME exporter, suggest realistic cost-overrun and saving percentages "
                "for a pessimistic and an optimistic budget scenario, by cost type (implementation, licences, training, oversight).",
         raw="Pessimistic: implementation +50%, licences +40%, training +10%. Optimistic: implementation -30%, licences -25%. "
             "'Industry studies show 70% of AI projects overrun budgets by 45%.'",
         edited="Pessimistic: implementation +25%, licences +20%, training +20%, oversight +10%; optimistic: -10% / -5% / 0% / 0%. "
                "All labelled 'assumed'.",
         why="The '70% / 45%' statistic had no source, so it was removed (ground rule). We raised training overrun (low digital skills means more "
             "sessions) and used milder swings because most of our spend is fixed-price SaaS."),
]


def doc():
    d = Doc("Finance Pack: Peer Annual-Report Summary", f"Prototype P3 - AI-drafted summary of {PEER} FY2024-25, verification table, budget scenarios",
            "Member 3 - Chief Financial Officer", "P3-Finance")
    d.callout(["**Why this peer:** S.P. Apparels is a listed knitwear exporter from the same Tiruppur cluster, selling to the same EU/UK "
               "buyers. Its annual report is public, so it is a fair benchmark for WeaveWorks (fictional).",
               "**How we worked:** (1) an AI tool drafted the summary; (2) the CFO checked every figure against the official Annual Report "
               "2024-25; (3) corrections are logged below. Rule: no AI number goes into a board paper unless it is verified."],
              kind="info", title="Summary")
    d.h1("1. Verified summary (FY2024-25)")
    d.table(["Topic", "What the annual report and results say (verified version)"], SUMMARY, widths=[3.0, 14.0], size=8.8, first_col_bold=True)
    d.h1("2. Verification table")
    d.table(["#", "Claim", "AI draft said", "Verified value", "Status", "Page (AR PDF)"],
            [(i, c, a, v, st, pg) for i, (c, a, v, s, st, pg) in enumerate(VERIFY, 1)],
            widths=[0.6, 3.0, 2.4, 5.4, 2.6, 3.0], size=7.8)
    n_corr = sum("CORRECTED" in v[4] for v in VERIFY)
    d.p(f"**Result:** {len(VERIFY)} claims checked against the official Annual Report 2024-25: {n_corr} corrected, 1 added, "
        f"{sum(v[4] == 'Verified' for v in VERIFY)} confirmed. **Error rate of the raw AI draft: {n_corr} of {len(VERIFY)} claims.** "
        "That is why every AI number has a named human owner.")
    d.callout("Every figure was checked against the official S.P. Apparels Annual Report 2024-25 (286-page PDF, s-p-apparels.com, Investors). "
              "Page numbers are PDF page numbers. The FTA signing date comes from government sources, since the report only calls it 'recent'.",
              kind="ok", title="Primary-source check complete")
    d.h1("3. What the benchmark means for WeaveWorks")
    d.table(["Metric", "S.P. Apparels FY25", "WeaveWorks", "So what"], BENCH, widths=[3.2, 3.2, 2.8, 7.8], size=8.5, first_col_bold=True)
    d.h1("4. Three-scenario AI budget (Rs 50 lakh, 18 months)")
    bs = facts.budget_summary()
    d.p(f"Base plan: **Rs {bs['total']:.2f} lakh** committed (use cases Rs {bs['uc_cost']:.2f} lakh + hidden costs Rs {bs['hidden']:.2f} lakh + "
        f"contingency Rs {bs['contingency']:.2f} lakh), leaving **Rs {bs['headroom']:.2f} lakh unallocated**, released only at Stage-gate 1 (M6). "
        "The workbook `Finance_Budget_and_Peer.xlsx` recalculates all three scenarios from the yellow input cells.")
    d.table(["Scenario", "Cost assumptions (assumed)", "What happens"], [
        ("Pessimistic", "Implementation +25%, licences/usage +20%, training +20%, oversight +10%",
         "Spend before contingency rises above Rs 50 lakh. Contingency is consumed and the stage-gate cut list applies: defer UC11 (Tamil voice SOP) to FY29 and trim UC13."),
        ("Base", "As planned", "Within budget; contingency kept for risks; unallocated reserve decided at M6."),
        ("Optimistic", "Implementation -10%, licences -5%", "Savings move to the H3 pilot (UC10 computer-vision QC) if Stage-gate 2 approves."),
    ], widths=[2.4, 6.6, 8.0], size=8.5, first_col_bold=True)
    d.h1("5. Prompt log: prompt, raw output, edited version")
    d.prompt_log(LOG)
    d.callout("Rejected: the AI's 'virtually debt-free' statement about the peer. It was confident, plausible and wrong (D/E about 0.30). "
              "If it had reached the board it would have undermined every other number in our pitch.", kind="ai", title="Rejected AI output (for the viva)")
    d.sources(["SPAL_AR25", "SPAL_FY25_RESULTS", "SPAL_CONCALL", "SPAL_BRSR24", "INDIA_UK_FTA"],
              extra=["Aggregator cross-checks: Screener.in, Business Standard capital-market news, Kotak Securities results pages (accessed Nov 2026)."])
    return d.save(P3 / "Peer_Annual_Report_Summary.docx")


def xlsx():
    wb = Workbook()
    # ------------------------------------------------ Assumptions
    a = wb.active
    a.title = "Assumptions"
    title(a, "3-scenario AI budget - assumptions", "Blue = input. Yellow = key assumption you may change. All values ASSUMED unless noted. Rs lakh.")
    header(a, 4, ["Cost driver", "Pessimistic", "Base", "Optimistic", "Note"])
    drivers = [("One-time implementation / set-up", 0.25, 0, -0.10, "Vendor overrun, data clean-up, integration surprises"),
               ("Monthly licences & usage", 0.20, 0, -0.05, "Seat/usage pricing changes, FX on USD-priced SaaS"),
               ("Training & change management", 0.20, 0, 0, "Low digital skills -> more sessions, Tamil translation"),
               ("Human oversight", 0.10, 0, 0, "Extra review time in first months"),
               ("Security, legal & audit", 0.0, 0, 0, "Fixed-fee engagement")]
    for i, (n, p, b, o, nt) in enumerate(drivers):
        r = 5 + i
        cell(a, r, 1, n, color="000000")
        for j, v in enumerate((p, b, o)):
            cell(a, r, 2 + j, v, fmt=PCT, fill=INPUT_FILL)
        cell(a, r, 5, nt, color="000000", wrap=True)
    cell(a, 11, 1, "Programme budget (Rs lakh)", bold=True, color="000000")
    cell(a, 11, 3, facts.BUDGET_LAKH, fmt=LAKH, fill=INPUT_FILL)
    cell(a, 12, 1, "Contingency (Rs lakh)", bold=True, color="000000")
    cell(a, 12, 3, facts.CONTINGENCY_LAKH, fmt=LAKH, fill=INPUT_FILL)
    widths(a, [36, 13, 13, 13, 52])

    # ------------------------------------------------ Budget
    b = wb.create_sheet("Budget_3_Scenarios")
    title(b, "AI programme budget by line, three scenarios (Rs lakh, 18 months)",
          "Use-case costs = one-time + monthly run x months live (from ROI model). Scenario = base x (1 + driver % from Assumptions).")
    header(b, 4, ["ID", "Budget line", "Included (1/0)", "One-time base", "Run / month base", "Months live",
                  "Base total", "Pessimistic", "Base", "Optimistic"])
    rows = [r for r in facts.roi_rows()]
    r0 = 5
    for i, rr in enumerate(rows):
        r = r0 + i
        cell(b, r, 1, rr["id"], color="000000")
        cell(b, r, 2, rr["name"], color="000000")
        cell(b, r, 3, 1 if rr["included"] else 0, fill=INPUT_FILL, fmt="0",
             comment="UC10 set to 0 after Shock E: deferred to fund UC17." if rr["id"] == "UC10" else None)
        cell(b, r, 4, rr["one_time"], fmt=LAKH)
        cell(b, r, 5, rr["run_pm"], fmt=LAKH)
        cell(b, r, 6, rr["months_live"], fmt="0")
        cell(b, r, 7, f"=C{r}*(D{r}+E{r}*F{r})", fmt=LAKH)
        cell(b, r, 8, f"=C{r}*(D{r}*(1+Assumptions!$B$5)+E{r}*F{r}*(1+Assumptions!$B$6))", fmt=LAKH)
        cell(b, r, 9, f"=C{r}*(D{r}*(1+Assumptions!$C$5)+E{r}*F{r}*(1+Assumptions!$C$6))", fmt=LAKH)
        cell(b, r, 10, f"=C{r}*(D{r}*(1+Assumptions!$D$5)+E{r}*F{r}*(1+Assumptions!$D$6))", fmt=LAKH)
    r_uc_end = r0 + len(rows) - 1
    r = r_uc_end + 1
    cell(b, r, 2, "Subtotal: use cases", bold=True, color="000000")
    for col, L in zip(range(7, 11), "GHIJ"):
        cell(b, r, col, f"=SUM({L}{r0}:{L}{r_uc_end})", fmt=LAKH, bold=True)
    r_sub = r
    hidden_rows = []
    for k, ((name, val), arow) in enumerate(zip(facts.HIDDEN_COSTS, (7, 8, 9))):
        r += 1
        hidden_rows.append(r)
        cell(b, r, 1, "HC", color="000000")
        cell(b, r, 2, name, color="000000", wrap=True)
        cell(b, r, 7, val, fmt=LAKH)
        for col, L in zip(range(8, 11), "BCD"):
            cell(b, r, col, f"=G{r}*(1+Assumptions!{L}${arow})", fmt=LAKH)
    r += 1
    cell(b, r, 2, "Total spend before contingency", bold=True, color="000000")
    for col, L in zip(range(7, 11), "GHIJ"):
        cell(b, r, col, f"={L}{r_sub}+SUM({L}{hidden_rows[0]}:{L}{hidden_rows[-1]})", fmt=LAKH, bold=True)
    r_spend = r
    r += 1
    cell(b, r, 2, "Contingency held (base) / consumed (pessimistic)", color="000000")
    cell(b, r, 7, "=Assumptions!$C$12", fmt=LAKH, color="008000")
    cell(b, r, 8, f"=MIN(Assumptions!$C$12,MAX(0,H{r_spend}-(Assumptions!$C$11-Assumptions!$C$12)))", fmt=LAKH)
    cell(b, r, 9, "=Assumptions!$C$12", fmt=LAKH, color="008000")
    cell(b, r, 10, "=Assumptions!$C$12", fmt=LAKH, color="008000")
    r_cont = r
    r += 1
    cell(b, r, 2, "Total committed (spend + contingency)", bold=True, color="000000", fill=SUB_FILL)
    cell(b, r, 7, f"=G{r_spend}+G{r_cont}", fmt=LAKH, bold=True, fill=SUB_FILL)
    cell(b, r, 8, f"=MAX(H{r_spend},H{r_spend}+Assumptions!$C$12-H{r_cont})", fmt=LAKH, bold=True, fill=SUB_FILL,
         comment="Pessimistic: contingency is consumed by overruns, so committed = max(spend, spend + unused contingency).")
    cell(b, r, 9, f"=I{r_spend}+I{r_cont}", fmt=LAKH, bold=True, fill=SUB_FILL)
    cell(b, r, 10, f"=J{r_spend}+J{r_cont}", fmt=LAKH, bold=True, fill=SUB_FILL)
    r_tot = r
    r += 1
    cell(b, r, 2, "Headroom vs Rs 50 lakh (negative = gap)", bold=True, color="000000")
    for col, L in zip(range(7, 11), "GHIJ"):
        cell(b, r, col, f"=Assumptions!$C$11-{L}{r_tot}", fmt=LAKH, bold=True)
    r_head = r
    r += 1
    cell(b, r, 2, "Status", bold=True, color="000000")
    for col, L in zip(range(7, 11), "GHIJ"):
        cell(b, r, col, f'=IF({L}{r_head}>=0,"Within budget","Apply cut list")', bold=True)
    r += 2
    note(b, r, "Stage-gate cut list if pessimistic (in order): 1) defer UC11 Tamil voice SOP assistant (frees its full cost); "
               "2) cut UC13 to JD checker only; 3) delay UC09 go-live by 2 months (frees 2 months' run cost). "
               "Never cut: training, human oversight, security/legal - these are the controls that make the rest safe.", span=10)
    widths(b, [6, 44, 10, 11, 11, 9, 11, 12, 11, 12])
    b.freeze_panes = "C5"

    # ------------------------------------------------ Verification
    v = wb.create_sheet("Verification")
    title(v, f"Peer annual-report verification - {PEER}", "Checked against the official Annual Report 2024-25 PDF (286 pages); page refs are PDF page numbers.")
    header(v, 4, ["#", "Claim", "AI draft said", "Verified value", "Source", "Status", "Page ref (AR 2024-25)", "Checked by"])
    for i, (c, ai, val, s, st, pg) in enumerate(VERIFY, 1):
        r = 4 + i
        for j, x in enumerate((i, c, ai, val, facts.SOURCES[s], st, pg)):
            cell(v, r, j + 1, x, wrap=True, color="000000")
        cell(v, r, 8, "Member 3 (CFO)", color="000000")
    widths(v, [4, 24, 20, 36, 40, 22, 14, 14])
    p = P3 / "Finance_Budget_and_Peer.xlsx"
    wb.save(p)
    return p


if __name__ == "__main__":
    print(xlsx())
    print(doc())
