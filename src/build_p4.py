"""P4 - Operations: demand forecast workbook (Excel formulas, same method as analysis.py), Power BI star-schema CSVs +
DAX variance measures, supply-disruption decision matrix, Ops pack (docx)."""
import json

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, BarChart
from openpyxl.utils import get_column_letter as L

import facts
from docx_lib import Doc, FIG, ROOT
from xlsx_lib import (INPUT_FILL, INT, PCT, SUB_FILL, cell, header, note, title, widths)

P4 = ROOT / "02_Build" / "P4_Operations"
PBI = P4 / "powerbi_data"
R = json.loads((ROOT / "data" / "results.json").read_text())
SHORT = {"Basic & fashion T-shirts": "Tshirts", "Polos": "Polos", "Kidswear": "Kidswear", "Loungewear": "Loungewear"}

# --------------------------------------------------------- decision matrix
CRITERIA = [("On-time delivery to buyers", 0.30), ("Cost / margin impact", 0.25),
            ("Compliance & traceability risk (GOTS chain of custody, due diligence)", 0.20),
            ("Buyer relationship & reputation", 0.15), ("Speed and ease of execution", 0.10)]
OPTIONS = [
    ("A. Wait for main spinner; ask buyers for 3-4 week extensions", [1, 4, 5, 1, 5]),
    ("B. Spot-buy from 2 pre-qualified alternate spinners at +12%", [4, 2, 4, 4, 3]),
    ("C. Move 30% of volume to a job-work partner that supplies its own fabric", [4, 2, 2, 3, 3]),
    ("D. Re-sequence with the forecast: protect Segment A/GOTS orders, partial-ship Segment B, air-freight only critical", [3, 3, 5, 3, 4]),
    ("E. Substitute yarn spec (different count/blend) with buyer approval", [3, 4, 1, 2, 2]),
    ("F. Hybrid B + D (recommended)", [4, 3, 4, 4, 3]),
]
DISRUPTION = ("Scenario (synthetic): a fire stops WeaveWorks' main cotton-yarn spinner (45% of yarn volume, GOTS-certified) for 6 weeks at the start "
              "of the Autumn/Winter peak. Spot yarn prices jump about 12%. At the same time, Red Sea rerouting adds 10-14 days to EU transit. "
              "Peak-month forecast already exceeds in-house capacity.")


def write_model(wb, prod, months, actual, train_bt, n_fwd):
    """One sheet per product line. Rows = t. Backtest block fits rows 1..train_bt; Final block fits all actuals."""
    ws = wb.create_sheet(f"FC_{SHORT[prod]}")
    N = len(actual)
    T = N + n_fwd
    title(ws, f"Forecast model - {prod}",
          "Classical decomposition: seasonal index from 2x12 centred moving-average ratios, x linear trend (FORECAST) on deseasonalised data. "
          f"Back-test fits t=1..{train_bt}, tests t={train_bt+1}..{N}. Final fits t=1..{N}, forecasts t={N+1}..{T}.")
    hdr = ["Month", "t", "Cal month", "Actual pcs",
           "BT: CMA12", "BT: ratio", "BT: SI", "BT: deseason", "BT: forecast",
           "Final: CMA12", "Final: ratio", "Final: SI", "Final: deseason", "Final: fitted/forecast"]
    header(ws, 4, hdr)
    r0 = 5
    rl = r0 + T - 1
    # SI tables at columns P..S
    header(ws, 4, ["Cal month", "BT raw SI", "BT SI", "Final raw SI", "Final SI"], col=16)
    for m in range(1, 13):
        r = 4 + m
        cell(ws, r, 16, m, fmt="0", color="000000")
        cell(ws, r, 17, f"=AVERAGEIFS($F${r0}:$F${rl},$C${r0}:$C${rl},P{r})", fmt="0.0000")
        cell(ws, r, 18, f"=Q{r}*12/SUM($Q$5:$Q$16)", fmt="0.0000")
        cell(ws, r, 19, f"=AVERAGEIFS($K${r0}:$K${rl},$C${r0}:$C${rl},P{r})", fmt="0.0000")
        cell(ws, r, 20, f"=S{r}*12/SUM($S$5:$S$16)", fmt="0.0000")
    bt_last = r0 + train_bt - 1
    fin_last = r0 + N - 1
    for i in range(T):
        r = r0 + i
        t = i + 1
        cell(ws, r, 1, months[i], color="000000")
        cell(ws, r, 2, t, fmt="0", color="000000")
        cell(ws, r, 3, f"=MOD(B{r}-1,12)+1", fmt="0")
        if i < N:
            cell(ws, r, 4, int(actual[i]), fmt=INT)
        else:
            cell(ws, r, 4, None)
        # backtest block
        ok_bt = 7 <= t <= train_bt - 6
        cell(ws, r, 5, f"=(AVERAGE(D{r-6}:D{r+5})+AVERAGE(D{r-5}:D{r+6}))/2" if ok_bt else "", fmt=INT)
        cell(ws, r, 6, f"=D{r}/E{r}" if ok_bt else "", fmt="0.000")
        cell(ws, r, 7, f"=INDEX($R$5:$R$16,C{r})", fmt="0.000")
        cell(ws, r, 8, f"=D{r}/G{r}" if t <= train_bt else "", fmt=INT)
        cell(ws, r, 9, f"=FORECAST(B{r},$H${r0}:$H${bt_last},$B${r0}:$B${bt_last})*G{r}" if train_bt < t <= N else "", fmt=INT)
        ok_f = 7 <= t <= N - 6
        cell(ws, r, 10, f"=(AVERAGE(D{r-6}:D{r+5})+AVERAGE(D{r-5}:D{r+6}))/2" if ok_f else "", fmt=INT)
        cell(ws, r, 11, f"=D{r}/J{r}" if ok_f else "", fmt="0.000")
        cell(ws, r, 12, f"=INDEX($T$5:$T$16,C{r})", fmt="0.000")
        cell(ws, r, 13, f"=D{r}/L{r}" if t <= N else "", fmt=INT)
        cell(ws, r, 14, f"=FORECAST(B{r},$M${r0}:$M${fin_last},$B${r0}:$B${fin_last})*L{r}", fmt=INT)
    widths(ws, [10, 5, 7, 11, 11, 9, 8, 11, 11, 11, 9, 8, 11, 13, 2, 9, 10, 9, 11, 9])
    ws.freeze_panes = "E5"
    return ws.title, r0, train_bt, N, T


def build_xlsx():
    o = pd.read_csv(ROOT / "data" / "orders_monthly.csv")
    piv = o.pivot(index="month", columns="product_line", values="shipped_pieces")[facts.PRODUCTS]
    months = [m[:7] for m in piv.index] + R["forecast"]["future_months"]
    train_bt = R["forecast"]["train_bt"]
    wb = Workbook()
    rd = wb.active
    rd.title = "Read_Me"
    title(rd, "WeaveWorks Exports - Demand forecast & supply-disruption decision (Prototype P4)",
          "All data SYNTHETIC (src/gen_data.py). Blue = input data; black = formula; yellow = assumption you may change.")
    lines = [
        "Sheets:",
        "  FC_<product>  - forecast model per product line (live formulas: seasonal index x linear trend).",
        "  Variance      - back-test Oct 2025-Sep 2026: actual vs forecast, variance, variance %, MAPE (same logic as the DAX measures).",
        "  Forward       - 12-month forward forecast by product and total, vs in-house capacity; chart.",
        "  Decision_Matrix - weighted scoring of six responses to a yarn-supply disruption.",
        "Power BI: load powerbi_data/*.csv and paste the measures in DAX_Measures.md (Forecast Variance, Variance %, MAPE).",
        "Method note: the Python version (src/analysis.py) produces identical numbers; it was used to check the formulas.",
    ]
    for i, t in enumerate(lines):
        rd.cell(4 + i, 1, t)
    widths(rd, [120])
    models = {}
    for p in facts.PRODUCTS:
        models[p] = write_model(wb, p, months, piv[p].to_numpy(), train_bt, 12)

    # ------------------------------------------------ Variance (backtest)
    v = wb.create_sheet("Variance", 1)
    title(v, "Back-test variance: actual vs forecast, Oct 2025 - Sep 2026",
          "Forecast Variance = Actual - Forecast; Variance % = Variance / Forecast; APE = |Variance| / Actual. Mirrors DAX measures.")
    hdr = ["Month"]
    for p in facts.PRODUCTS:
        s = SHORT[p]
        hdr += [f"{s} actual", f"{s} fcst", f"{s} var", f"{s} APE"]
    hdr += ["Total actual", "Total fcst", "Total var", "Total var %", "Total APE"]
    header(v, 4, hdr, height=42)
    N = models[facts.PRODUCTS[0]][3]
    rows = list(range(5, 5 + (N - train_bt)))
    for k, r in enumerate(rows):
        src_r = 5 + train_bt + k
        sh0 = models[facts.PRODUCTS[0]][0]
        cell(v, r, 1, f"={sh0}!A{src_r}", color="008000")
        for j, p in enumerate(facts.PRODUCTS):
            sh = models[p][0]
            c = 2 + j * 4
            cell(v, r, c, f"={sh}!D{src_r}", fmt=INT, color="008000")
            cell(v, r, c + 1, f"={sh}!I{src_r}", fmt=INT, color="008000")
            cell(v, r, c + 2, f"={L(c)}{r}-{L(c+1)}{r}", fmt=INT)
            cell(v, r, c + 3, f"=ABS({L(c+2)}{r})/{L(c)}{r}", fmt=PCT)
        ca = 2 + 4 * len(facts.PRODUCTS)
        cell(v, r, ca, "=" + "+".join(f"{L(2 + j*4)}{r}" for j in range(4)), fmt=INT)
        cell(v, r, ca + 1, "=" + "+".join(f"{L(3 + j*4)}{r}" for j in range(4)), fmt=INT)
        cell(v, r, ca + 2, f"={L(ca)}{r}-{L(ca+1)}{r}", fmt=INT)
        cell(v, r, ca + 3, f"={L(ca+2)}{r}/{L(ca+1)}{r}", fmt=PCT)
        cell(v, r, ca + 4, f"=ABS({L(ca+2)}{r})/{L(ca)}{r}", fmt=PCT)
    rs = rows[-1] + 1
    cell(v, rs, 1, "Sum / MAPE", bold=True, color="000000", fill=SUB_FILL)
    for j in range(4):
        c = 2 + j * 4
        for k in range(3):
            cell(v, rs, c + k, f"=SUM({L(c+k)}{rows[0]}:{L(c+k)}{rows[-1]})", fmt=INT, bold=True, fill=SUB_FILL)
        cell(v, rs, c + 3, f"=AVERAGE({L(c+3)}{rows[0]}:{L(c+3)}{rows[-1]})", fmt=PCT, bold=True, fill=SUB_FILL)
    ca = 2 + 16
    for k in range(3):
        cell(v, rs, ca + k, f"=SUM({L(ca+k)}{rows[0]}:{L(ca+k)}{rows[-1]})", fmt=INT, bold=True, fill=SUB_FILL)
    cell(v, rs, ca + 3, f"={L(ca+2)}{rs}/{L(ca+1)}{rs}", fmt=PCT, bold=True, fill=SUB_FILL)
    cell(v, rs, ca + 4, f"=AVERAGE({L(ca+4)}{rows[0]}:{L(ca+4)}{rows[-1]})", fmt=PCT, bold=True, fill=SUB_FILL)
    cell(v, rs + 2, 1, "Benchmark: current manual forecast MAPE (assumed)", color="000000")
    cell(v, rs + 2, 6, facts.CO["forecast_mape_now"], fmt=PCT, fill=INPUT_FILL)
    v.merge_cells(start_row=rs + 2, start_column=1, end_row=rs + 2, end_column=5)
    widths(v, [10] + [10] * 21)
    v.freeze_panes = "B5"

    # ------------------------------------------------ Forward
    f = wb.create_sheet("Forward", 2)
    title(f, "12-month forward forecast vs in-house capacity (pieces)", "Capacity is an assumption (yellow). Gap = months where outsourcing/overtime must be booked early.")
    cell(f, 3, 1, "In-house capacity / month", color="000000")
    cell(f, 3, 3, facts.CO["capacity_pcs_per_month"], fmt=INT, fill=INPUT_FILL)
    header(f, 5, ["Month"] + [SHORT[p] for p in facts.PRODUCTS] + ["Total", "Capacity", "Gap (pcs)"])
    for k in range(12):
        r = 6 + k
        src_r = 5 + N + k
        cell(f, r, 1, f"={models[facts.PRODUCTS[0]][0]}!A{src_r}", color="008000")
        for j, p in enumerate(facts.PRODUCTS):
            cell(f, r, 2 + j, f"={models[p][0]}!N{src_r}", fmt=INT, color="008000")
        cell(f, r, 6, f"=SUM(B{r}:E{r})", fmt=INT)
        cell(f, r, 7, "=$C$3", fmt=INT)
        cell(f, r, 8, f"=MAX(0,F{r}-G{r})", fmt=INT)
    cell(f, 18, 1, "12-month total", bold=True, color="000000", fill=SUB_FILL)
    for c in range(2, 9):
        cell(f, 18, c, f"=SUM({L(c)}6:{L(c)}17)", fmt=INT, bold=True, fill=SUB_FILL)
    cell(f, 19, 1, "Growth vs last 12 months actual", color="000000")
    last12 = "+".join(f"SUM({m[0]}!D{5 + N - 12}:D{4 + N})" for m in models.values())
    cell(f, 19, 6, f"=F18/({last12})-1", fmt=PCT)
    cell(f, 20, 1, "Months above capacity", color="000000")
    cell(f, 20, 6, "=COUNTIF(H6:H17,\">0\")", fmt="0")
    ch = LineChart()
    ch.title = "Forward forecast vs capacity"
    ch.y_axis.title = "Pieces"
    ch.add_data(Reference(f, min_col=6, min_row=5, max_row=17), titles_from_data=True)
    ch.add_data(Reference(f, min_col=7, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(f, min_col=1, min_row=6, max_row=17))
    ch.height, ch.width = 8, 18
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.series[0].graphicalProperties.line.solidFill = "2A78D6"
    ch.series[1].graphicalProperties.line.solidFill = "EB6834"
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    f.add_chart(ch, "J5")
    widths(f, [10, 11, 11, 11, 11, 12, 11, 11])

    # ------------------------------------------------ Decision matrix
    dm = wb.create_sheet("Decision_Matrix", 3)
    title(dm, "Supply-disruption decision matrix (weighted scoring, 1 = poor, 5 = best)")
    note(dm, 2, DISRUPTION, span=9)
    dm.row_dimensions[2].height = 52
    header(dm, 4, ["Criterion", "Weight"] + [o[0].split(".")[0] for o in OPTIONS])
    for i, (c, w) in enumerate(CRITERIA):
        r = 5 + i
        cell(dm, r, 1, c, wrap=True, color="000000")
        cell(dm, r, 2, w, fmt=PCT, fill=INPUT_FILL)
        for j, (_, sc) in enumerate(OPTIONS):
            cell(dm, r, 3 + j, sc[i], fmt="0", fill=INPUT_FILL)
    rw = 5 + len(CRITERIA)
    cell(dm, rw, 1, "Weights check (must be 100%)", italic=True, color="000000")
    cell(dm, rw, 2, f"=SUM(B5:B{rw-1})", fmt=PCT)
    cell(dm, rw + 1, 1, "Weighted score (max 5)", bold=True, color="000000", fill=SUB_FILL)
    for j in range(len(OPTIONS)):
        c = L(3 + j)
        cell(dm, rw + 1, 3 + j, f"=SUMPRODUCT($B$5:$B${rw-1},{c}5:{c}{rw-1})", fmt="0.00", bold=True, fill=SUB_FILL)
    cell(dm, rw + 2, 1, "Rank", bold=True, color="000000")
    for j in range(len(OPTIONS)):
        c = L(3 + j)
        cell(dm, rw + 2, 3 + j, f"=RANK({c}{rw+1},$C${rw+1}:${L(2+len(OPTIONS))}${rw+1})", fmt="0", bold=True)
    cell(dm, rw + 3, 1, "Best option", bold=True, color="000000")
    cell(dm, rw + 3, 3, f"=INDEX($C$4:${L(2+len(OPTIONS))}$4,MATCH(MAX($C${rw+1}:${L(2+len(OPTIONS))}${rw+1}),$C${rw+1}:${L(2+len(OPTIONS))}${rw+1},0))", bold=True)
    r = rw + 5
    header(dm, r, ["Option", "Description"])
    for j, (o, _) in enumerate(OPTIONS):
        cell(dm, r + 1 + j, 1, o.split(".")[0], color="000000")
        cell(dm, r + 1 + j, 2, o.split(". ", 1)[1], color="000000")
        dm.merge_cells(start_row=r + 1 + j, start_column=2, end_row=r + 1 + j, end_column=8)
    widths(dm, [44, 9, 8, 8, 8, 8, 8, 8])
    p = P4 / "Demand_Forecast_and_Supply_Decision.xlsx"
    wb.move_sheet("Read_Me", offset=-10)
    wb.save(p)
    return p


def powerbi():
    PBI.mkdir(parents=True, exist_ok=True)
    o = pd.read_csv(ROOT / "data" / "orders_monthly.csv")
    o.rename(columns={"month": "month_start"}).to_csv(PBI / "fact_shipments.csv", index=False)
    rows = []
    fc = R["forecast"]
    bt_months = fc["months"][fc["train_bt"]:]
    for p, d in fc["lines"].items():
        for m, val in zip(bt_months, d["backtest"]):
            rows.append(dict(month_start=m + "-01", product_line=p, forecast_type="Backtest", forecast_pieces=round(val)))
        for m, val in zip(fc["future_months"], d["forward"]):
            rows.append(dict(month_start=m + "-01", product_line=p, forecast_type="Forward", forecast_pieces=round(val)))
    pd.DataFrame(rows).to_csv(PBI / "fact_forecast.csv", index=False)
    pd.DataFrame(dict(product_line=facts.PRODUCTS, category=["Tops", "Tops", "Kids", "Lounge"],
                      gots_share_assumed=[0.35, 0.20, 0.60, 0.40])).to_csv(PBI / "dim_product.csv", index=False)
    dates = pd.date_range("2023-01-01", "2027-09-01", freq="MS")
    pd.DataFrame(dict(month_start=dates.strftime("%Y-%m-%d"), year=dates.year, month=dates.month,
                      month_name=dates.strftime("%b"), fy=[f"FY{(d.year + (1 if d.month >= 4 else 0)) % 100:02d}" for d in dates])
                 ).to_csv(PBI / "dim_date.csv", index=False)
    (P4 / "DAX_Measures.md").write_text(DAX)


DAX = r"""# Power BI build: forecast variance with DAX (Prototype P4)

## 1. Load data (Get Data > Text/CSV) from `powerbi_data/`
- `fact_shipments.csv` (month_start, product_line, shipped_pieces)
- `fact_forecast.csv` (month_start, product_line, forecast_type = Backtest | Forward, forecast_pieces)
- `dim_product.csv`, `dim_date.csv` (mark dim_date as Date table on `month_start`; set month_start type = Date)

## 2. Relationships (Model view, single direction, many-to-one)
- fact_shipments[month_start] -> dim_date[month_start];  fact_shipments[product_line] -> dim_product[product_line]
- fact_forecast[month_start] -> dim_date[month_start];   fact_forecast[product_line] -> dim_product[product_line]

## 3. Measures (New measure, paste one at a time)
```DAX
Actual Pieces = SUM ( fact_shipments[shipped_pieces] )

Forecast Pieces =
CALCULATE ( SUM ( fact_forecast[forecast_pieces] ), fact_forecast[forecast_type] = "Backtest" )

Forward Forecast Pieces =
CALCULATE ( SUM ( fact_forecast[forecast_pieces] ), fact_forecast[forecast_type] = "Forward" )

-- The required DAX variance measure
Forecast Variance =
IF ( NOT ISBLANK ( [Forecast Pieces] ), [Actual Pieces] - [Forecast Pieces] )

Forecast Variance % = DIVIDE ( [Forecast Variance], [Forecast Pieces] )

Abs % Error = DIVIDE ( ABS ( [Forecast Variance] ), [Actual Pieces] )

MAPE =
AVERAGEX (
    FILTER (
        CROSSJOIN ( VALUES ( dim_date[month_start] ), VALUES ( dim_product[product_line] ) ),
        NOT ISBLANK ( [Forecast Pieces] ) && NOT ISBLANK ( [Actual Pieces] )
    ),
    [Abs % Error]
)

YoY Variance =
[Actual Pieces] - CALCULATE ( [Actual Pieces], SAMEPERIODLASTYEAR ( dim_date[month_start] ) )

Capacity Gap =
VAR Cap = 360000   -- assumed in-house capacity per month
RETURN MAX ( 0, SUMX ( VALUES ( dim_date[month_start] ), [Forward Forecast Pieces] - Cap ) )
```

## 4. Report page (one page is enough for the demo)
1. Line chart: X = dim_date[month_start]; Y = Actual Pieces, Forecast Pieces, Forward Forecast Pieces.
2. Clustered bar: Y = product_line; X = Forecast Variance (conditional colour: red if negative).
3. Cards: MAPE, Forecast Variance %, Capacity Gap.
4. Slicer: dim_product[product_line]; slicer: dim_date[fy].

## 5. Check your numbers
With all products selected and the back-test window (Oct 2025 - Sep 2026), Forecast Variance % should equal the Excel
`Variance` sheet total variance % and MAPE should equal the average of the product-level APEs across all 48 month-product cells.
"""


def doc():
    fc = R["forecast"]
    d = Doc("Operations Pack: Demand Forecast & Supply Decision", "Prototype P4 - Ops stream: forecast in Excel/Power BI with DAX variance measure, decision matrix for a supply disruption",
            "Member 4 - Chief Operating Officer", "P4-Operations")
    d.callout([f"**Forecast:** a seasonal-index x trend model on 45 months of synthetic shipments. In the 12-month back-test it was within "
               f"**{fc['total']['mape']:.1%} MAPE** at total level (product lines {min(v['mape'] for v in fc['lines'].values()):.1%}-"
               f"{max(v['mape'] for v in fc['lines'].values()):.1%}) vs an assumed {facts.CO['forecast_mape_now']:.0%} for today's manual forecast.",
               f"**Forward view:** {fc['total']['forward_sum']/1e6:.2f} million pieces over Oct 2026-Sep 2027 "
               f"(+{fc['total']['forward_sum']/sum(fc['total']['actual'][-12:]) - 1:.1%}). Several peak months exceed the assumed 360k/month in-house capacity, so job-work must be booked early.",
               "**Decision:** for a 6-week loss of the main yarn spinner, the weighted matrix recommends **F: spot-buy from pre-qualified spinners + re-sequence orders by forecast** (B + D)."],
              kind="info", title="Summary")
    d.h1("1. Why Operations (not HR) as the 4th prototype")
    d.p("EU buyers rank on-time delivery and compliance above price for repeat orders (team assumption from merchandiser interviews, to be "
        "validated). Forecast error causes three costs at WeaveWorks: yarn over-buying, last-minute job-work at premium rates, and air freight "
        "when shipments slip. The HR stream (bias-checked JDs and screening) is in the roadmap at H3 (UC13).")
    d.h1("2. Forecast method")
    d.numbered([
        "Data: monthly shipped pieces, Jan 2023 - Sep 2026, four product lines (synthetic, `data/orders_monthly.csv`).",
        "Seasonality: 2x12 centred moving average; ratio of actual to CMA; average ratio per calendar month, normalised to 12 (seasonal index, SI).",
        "Trend: linear regression (Excel FORECAST) on deseasonalised data (actual / SI).",
        "Forecast = trend x SI. Back-test: fit to Sep 2025, forecast Oct 2025 - Sep 2026, compare with actuals.",
        "Same method in Excel formulas (`Demand_Forecast_and_Supply_Decision.xlsx`) and Python (`src/analysis.py`); numbers reconcile.",
    ])
    d.image(FIG / "forecast_total.png", caption="Figure 1. Actual vs back-test and forward forecast (synthetic data).")
    rows = [(p, f"{v['mape']:.1%}", f"{v['bias']:+.1%}", f"{sum(v['forward'])/1e3:,.0f}k") for p, v in fc["lines"].items()]
    rows.append(("Total", f"{fc['total']['mape']:.1%}", f"{fc['total']['variance_pct']:+.1%}", f"{fc['total']['forward_sum']/1e3:,.0f}k"))
    d.table(["Product line", "Back-test MAPE", "Variance % (bias)", "Forward 12-month forecast"], rows, widths=[5.5, 3.5, 3.8, 4.2], size=9)
    d.callout("The synthetic series is smooth by design, so 3-8% MAPE flatters the model. On real ERP data, expect 10-15%; that would still "
              "halve today's assumed error. Do not quote the 3% figure to the board as a promise.", kind="warn", title="Honest caveat")
    d.h1("3. Power BI and the DAX variance measure")
    d.p("`powerbi_data/` holds a star schema (two fact tables, product and date dimensions). `DAX_Measures.md` gives step-by-step build "
        "instructions and the measures. The required variance measure is:")
    d.p("**Forecast Variance = IF ( NOT ISBLANK ( [Forecast Pieces] ), [Actual Pieces] - [Forecast Pieces] )**  and  "
        "**Forecast Variance % = DIVIDE ( [Forecast Variance], [Forecast Pieces] )**, plus a MAPE measure using AVERAGEX over month x product.", size=9.5)
    d.p("The Excel `Variance` sheet calculates the same quantities with formulas, so the Power BI card values can be checked line by line.")
    d.h1("4. Supply-disruption decision matrix")
    d.p(DISRUPTION)
    score = lambda sc: sum(w * s for (_, w), s in zip(CRITERIA, sc))
    d.table(["Option", "Delivery 30%", "Cost 25%", "Compliance 20%", "Relationship 15%", "Ease 10%", "Score"],
            [(o, *sc, f"{score(sc):.2f}") for o, sc in OPTIONS], widths=[7.4, 1.7, 1.5, 1.9, 1.9, 1.4, 1.2], size=8)
    best = max(OPTIONS, key=lambda x: score(x[1]))
    d.p(f"**Recommendation: {best[0]}** (score {score(best[1]):.2f}/5). Option E (substitute yarn) looks cheap but breaks the GOTS chain of "
        "custody, which would put Segment A buyers' compliance at risk, so it scores 1 on compliance. Option A protects margin but damages "
        "on-time delivery during the season that matters most.")
    d.bullets([
        "**Day 1-2:** activate the two pre-qualified spinners (ask for GOTS transaction certificates before dispatch); freeze new Segment B order intake for 2 weeks.",
        "**Day 2-3:** re-run the forecast by buyer/style; protect Segment A and GOTS orders; offer Segment B buyers partial shipments with no price change.",
        "**Day 3-5:** air-freight only orders with a contractual penalty above the air-freight cost (finance check); inform buyers proactively (the support Agent's at-risk flag).",
        "**After:** dual-source policy, with no spinner above 35% of yarn volume, and a quarterly supplier-risk review in the risk register.",
    ])
    w2 = [0.30, 0.45, 0.0, 0.15, 0.10]
    alt = sorted(OPTIONS, key=lambda x: -sum(a * b for a, b in zip(w2, x[1])))
    rank_e = [o[0][0] for o in alt].index("E") + 1
    d.callout(f"If the compliance weight is set to 0% (moved to cost), {alt[0][0][0]} is still first, but option E "
              f"(substitute yarn spec) jumps from last to rank {rank_e}. The risky option only looks attractive when EU due diligence "
              "is ignored. Test other weights in the yellow cells.", kind="info", title="Sensitivity")
    d.h1("5. Prompt log: prompt, raw output, edited version")
    d.prompt_log(OPS_LOG)
    d.callout("Rejected: the AI's first forecast suggestion, a 12-month moving average with no seasonality. It would have under-forecast every "
              "SS and AW peak by 15-20% and hidden the capacity gap. We kept the AI's suggestion only as the naive baseline.",
              kind="ai", title="Rejected AI output (for the viva)")
    d.sources([], extra=["Red Sea / Suez rerouting adding 10-14 days to Asia-Europe transit (2024-25) is widely reported in shipping-industry news; "
                         "treat the exact days as assumed."])
    return d.save(P4 / "Operations_Pack.docx")


OPS_LOG = [
    dict(title="Choose a forecasting method for Excel",
         prompt="We have 45 months of monthly shipment data for 4 knitwear product lines with strong seasonality (SS and AW peaks). Suggest a "
                "forecasting method that a COO can build and explain in Excel without add-ins, and how to back-test it.",
         raw="'Use a 12-month moving average and project it forward. Alternatively use FORECAST.ETS which handles everything automatically.'",
         edited="Classical decomposition (CMA seasonal index x linear trend) with an explicit 12-month back-test; FORECAST.ETS kept as an optional comparison.",
         why="A plain moving average ignores seasonality. FORECAST.ETS is a black box the COO cannot explain in the viva or audit in the panel Q&A."),
    dict(title="Write the DAX variance measure",
         prompt="Write DAX measures for Actual, Forecast (only forecast_type = Backtest), Forecast Variance, Variance % and MAPE across month and product.",
         raw="'Forecast Variance = SUM(fact_shipments[shipped_pieces]) - SUM(fact_forecast[forecast_pieces])' and "
             "'MAPE = AVERAGE(ABS(fact_shipments[shipped_pieces]-fact_forecast[forecast_pieces]))'",
         edited="Variance uses the Backtest-filtered measure and returns blank outside the back-test window; MAPE uses AVERAGEX over a CROSSJOIN of month x product.",
         why="The raw version mixed Backtest and Forward rows (inflating forecast) and its MAPE is invalid DAX (AVERAGE cannot take an expression across two tables)."),
    dict(title="Score the disruption options",
         prompt="Our main GOTS-certified yarn spinner is down for 6 weeks at the AW peak. List response options and score them 1-5 on delivery, cost, "
                "compliance, relationships and ease, with weights.",
         raw="Ranked 'substitute a cheaper non-organic yarn and relabel later' as the best option (score 4.4) for cost and speed.",
         edited="Option E rewritten as 'substitute yarn spec with buyer approval' and scored 1 on compliance; recommendation is hybrid B + D.",
         why="'Relabel later' would be mislabelling organic products: a GOTS violation and potentially fraud. The model optimised the numbers and ignored the rule."),
]

if __name__ == "__main__":
    print(build_xlsx())
    powerbi()
    print(doc())
