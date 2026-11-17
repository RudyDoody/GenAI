# Power BI build: forecast variance with DAX (Prototype P4)

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
