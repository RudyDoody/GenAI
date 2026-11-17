"""Phase 4 - ROI model for every proposed use case: pessimistic / base / optimistic, break-even for the key driver, payback."""
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

import facts
from docx_lib import ROOT
from xlsx_lib import (INPUT_FILL, LAKH, PCT, SUB_FILL, cell, header, note, title, widths)

OUT = ROOT / "04_Decide" / "ROI_Model.xlsx"


def build():
    wb = Workbook()
    inp = wb.active
    inp.title = "Inputs"
    title(inp, "ROI model - global inputs", "Blue = input; yellow = key assumption; black = formula; green = cross-sheet link. Rs lakh unless stated.")
    rows = [("Scenario shown on Programme sheet", "Base", None, "Pick Pessimistic / Base / Optimistic"),
            ("Programme horizon (months)", facts.HORIZON_MONTHS, "0", "Brief: 18 months"),
            ("Programme budget (Rs lakh)", facts.BUDGET_LAKH, LAKH, "Brief: Rs 50 lakh"),
            ("Contingency (Rs lakh)", facts.CONTINGENCY_LAKH, LAKH, "About 9% of programme (assumed)")]
    for k, (n, h, _), in enumerate(facts.HIDDEN_COSTS and []):
        pass
    for i, (n, v, fmt, nt) in enumerate(rows):
        r = 4 + i
        cell(inp, r, 1, n, color="000000")
        cell(inp, r, 2, v, fmt=fmt, fill=INPUT_FILL, color="0000FF")
        cell(inp, r, 3, nt, color="000000")
    dv = DataValidation(type="list", formula1='"Pessimistic,Base,Optimistic"', allow_blank=False)
    inp.add_data_validation(dv)
    dv.add("B4")
    header(inp, 9, ["Hidden cost (shared, not in use-case lines)", "Rs lakh", "Why it is needed"])
    why = ["Low digital skills: Tamil-language sessions, AI champions, ADKAR reinforcement",
           "Human-in-the-loop review of AI outputs; part-time AI governance lead",
           "DPDP/GDPR review, vendor due diligence, annual AI audit, security testing"]
    for i, ((n, v), w) in enumerate(zip(facts.HIDDEN_COSTS, why)):
        r = 10 + i
        cell(inp, r, 1, n, color="000000", wrap=True)
        cell(inp, r, 2, v, fmt=LAKH, fill=INPUT_FILL)
        cell(inp, r, 3, w, color="000000", wrap=True)
    widths(inp, [52, 14, 70])

    u = wb.create_sheet("ROI_by_UseCase")
    title(u, "ROI by use case - 18-month window, three scenarios",
          "Benefit = annual driver volume x unit value (Rs lakh per unit) x months live / 12.  Cost = one-time + run/month x months live.  "
          "Break-even driver = annual volume needed for benefit = cost.  All driver volumes and unit values are ASSUMED.")
    hdr = ["ID", "Use case", "Incl. (1/0)", "Key driver", "Unit value (Rs lakh / unit)", "Pess. volume / yr", "Base volume / yr",
           "Opt. volume / yr", "Go-live month", "Months live", "One-time cost", "Run cost / month", "18-m cost",
           "Benefit pess.", "Benefit base", "Benefit opt.", "ROI pess.", "ROI base", "ROI opt.",
           "Break-even driver / yr", "Break-even as % of base", "Payback month (base)", "Selected-scenario benefit"]
    header(u, 4, hdr, height=48)
    rr = facts.roi_rows()
    r0 = 5
    for i, x in enumerate(rr):
        r = r0 + i
        cell(u, r, 1, x["id"], color="000000")
        cell(u, r, 2, x["name"], color="000000", wrap=True)
        cell(u, r, 3, 1 if x["included"] else 0, fmt="0", fill=INPUT_FILL,
             comment="Set to 0 after Shock E: deferred to fund UC17." if not x["included"] else None)
        cell(u, r, 4, f"{x['driver']} ({x['unit']})", color="000000", wrap=True)
        cell(u, r, 5, x["unit_value"], fmt="0.0000", fill=INPUT_FILL)
        for j, vol in enumerate(x["vols"]):
            cell(u, r, 6 + j, vol, fmt="#,##0.0", fill=INPUT_FILL if j == 1 else None)
        cell(u, r, 9, x["golive"], fmt="0", fill=INPUT_FILL)
        cell(u, r, 10, f"=Inputs!$B$5-I{r}+1", fmt="0")
        cell(u, r, 11, x["one_time"], fmt=LAKH)
        cell(u, r, 12, x["run_pm"], fmt=LAKH)
        cell(u, r, 13, f"=K{r}+L{r}*J{r}", fmt=LAKH)
        for j, col in enumerate("FGH"):
            cell(u, r, 14 + j, f"={col}{r}*$E{r}*$J{r}/12", fmt=LAKH)
        for j, col in enumerate("NOP"):
            cell(u, r, 17 + j, f"=IF($M{r}=0,0,({col}{r}-$M{r})/$M{r})", fmt=PCT)
        cell(u, r, 20, f"=M{r}/(E{r}*J{r}/12)", fmt="#,##0.00")
        cell(u, r, 21, f"=T{r}/G{r}", fmt=PCT)
        cell(u, r, 22, f'=IF(G{r}*E{r}/12-L{r}<=0,"Never",IF(I{r}+ROUNDUP(K{r}/(G{r}*E{r}/12-L{r}),0)-1>Inputs!$B$5,"After M18",'
                       f'I{r}+ROUNDUP(K{r}/(G{r}*E{r}/12-L{r}),0)-1))', fmt="0")
        cell(u, r, 23, f'=INDEX(N{r}:P{r},MATCH(Inputs!$B$4,{{"Pessimistic","Base","Optimistic"}},0))', fmt=LAKH)
    rl = r0 + len(rr) - 1
    rt = rl + 1
    cell(u, rt, 2, "Total - included use cases", bold=True, color="000000", fill=SUB_FILL)
    for col in range(11, 17):
        L = "ABCDEFGHIJKLMNOPQRSTUVW"[col - 1]
        if col in (11,):
            f = f"=SUMPRODUCT($C${r0}:$C${rl},K{r0}:K{rl})"
        elif col == 12:
            f = f"=SUMPRODUCT($C${r0}:$C${rl},L{r0}:L{rl})"
        else:
            f = f"=SUMPRODUCT($C${r0}:$C${rl},{L}{r0}:{L}{rl})"
        cell(u, rt, col, f, fmt=LAKH, bold=True, fill=SUB_FILL)
    for j, col in enumerate("NOP"):
        cell(u, rt, 17 + j, f"=({col}{rt}-$M{rt})/$M{rt}", fmt=PCT, bold=True, fill=SUB_FILL)
    cell(u, rt, 23, f"=SUMPRODUCT($C${r0}:$C${rl},W{r0}:W{rl})", fmt=LAKH, bold=True, fill=SUB_FILL)
    note(u, rt + 2, "Reading the break-even column: e.g. the support Agent breaks even inside 18 months if it saves the break-even number of "
                    "merchandiser hours per year; compare with the base volume. A break-even above 100% of base means the use case does not pay "
                    "back in the window on base assumptions (it may still be strategic). ROI here excludes shared hidden costs, which are "
                    "added on the Programme sheet.", span=14)
    widths(u, [6, 30, 7, 30, 10, 9, 9, 9, 8, 8, 9, 9, 9, 9, 9, 9, 8, 8, 8, 11, 10, 10, 11])
    u.freeze_panes = "C5"
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Base-case ROI by use case (18 months, direct cost)"
    ch.add_data(Reference(u, min_col=18, min_row=4, max_row=rl), titles_from_data=True)
    ch.set_categories(Reference(u, min_col=1, min_row=5, max_row=rl))
    ch.height, ch.width = 9, 16
    ch.legend = None
    ch.roundedCorners = False
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = "0%"
    ch.series[0].graphicalProperties.solidFill = "0F7C78"
    u.add_chart(ch, f"B{rt + 5}")

    p = wb.create_sheet("Programme", 1)
    title(p, "Programme view - all costs including hidden costs", "This is the number the board approves.")
    header(p, 4, ["Line", "Pessimistic", "Base", "Optimistic", "Note"])
    U = "ROI_by_UseCase"
    lines = [
        ("Use-case costs (direct)", f"={U}!M{rt}", f"={U}!M{rt}", f"={U}!M{rt}", "Same cost in all three: scenarios vary benefits (cost overruns: see Finance_Budget_and_Peer.xlsx)"),
        ("Hidden costs: training, oversight, security/legal", "=SUM(Inputs!$B$10:$B$12)", "=SUM(Inputs!$B$10:$B$12)", "=SUM(Inputs!$B$10:$B$12)", "Brief: include hidden costs"),
        ("Contingency", "=Inputs!$B$7", "=Inputs!$B$7", "=Inputs!$B$7", ""),
        ("Total programme cost", "=SUM(B5:B7)", "=SUM(C5:C7)", "=SUM(D5:D7)", ""),
        ("Headroom vs budget", "=Inputs!$B$6-B8", "=Inputs!$B$6-C8", "=Inputs!$B$6-D8", "Unallocated reserve, released at Stage-gate 1"),
        ("18-month benefits", f"={U}!N{rt}", f"={U}!O{rt}", f"={U}!P{rt}", "Sum of included use cases"),
        ("Net benefit", "=B10-B8", "=C10-C8", "=D10-D8", ""),
        ("Programme ROI", "=B11/B8", "=C11/C8", "=D11/D8", ""),
        ("Benefit realisation needed to break even", "=B8/C10", "=C8/C10", "=D8/C10", "Share of BASE benefits we must deliver to recover every rupee in 18 months"),
    ]
    for i, (n, a, b, c, nt) in enumerate(lines):
        r = 5 + i
        bold = n in ("Total programme cost", "Programme ROI", "Net benefit")
        cell(p, r, 1, n, bold=bold, color="000000")
        for j, f in enumerate((a, b, c)):
            cell(p, r, 2 + j, f, fmt=PCT if "ROI" in n or "realisation" in n else LAKH, bold=bold,
                 color="008000" if f.startswith(f"={U}") or f.startswith("=Inputs") or f.startswith("=SUM(Inputs") else None)
        cell(p, r, 5, nt, color="000000", wrap=True)
    cell(p, 15, 1, "Selected scenario (Inputs!B4)", bold=True, color="000000")
    cell(p, 15, 2, "=Inputs!B4", color="008000", bold=True)
    cell(p, 16, 1, "Selected-scenario benefit", color="000000")
    cell(p, 16, 2, f"={U}!W{rt}", fmt=LAKH, color="008000")
    cell(p, 17, 1, "Selected-scenario programme ROI", color="000000")
    cell(p, 17, 2, "=(B16-C8)/C8", fmt=PCT)
    header(p, 19, ["Sensitivity: % of base benefits delivered", "Benefit", "Net", "ROI"])
    for i, pct in enumerate([0.3, 0.4, 0.5, 0.6, 0.8, 1.0, 1.2]):
        r = 20 + i
        cell(p, r, 1, pct, fmt="0%", fill=INPUT_FILL)
        cell(p, r, 2, f"=A{r}*$C$10", fmt=LAKH)
        cell(p, r, 3, f"=B{r}-$C$8", fmt=LAKH)
        cell(p, r, 4, f"=C{r}/$C$8", fmt=PCT)
    widths(p, [46, 14, 14, 14, 60])
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
