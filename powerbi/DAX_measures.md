# Power BI Dashboard – Build Guide (about 30–40 min)

## 1. Load data
1. Power BI Desktop → **Get Data → Text/CSV** → `data/cleaned_sales_data.csv` → **Load**.
2. Check types in Data view: `Order_Date` = Date, `Sales`/`Cost`/`Profit`/`Unit_Price` = Fixed decimal, `Discount` = Percentage, `Quantity` = Whole number.
3. Rename the table to `Sales`.

## 2. Date table (Modeling → New table)
```DAX
DateTable =
ADDCOLUMNS(
    CALENDAR(DATE(2025,1,1), DATE(2026,5,31)),
    "Year", YEAR([Date]),
    "MonthNum", MONTH([Date]),
    "Month", FORMAT([Date], "MMM"),
    "MonthYear", FORMAT([Date], "MMM YYYY"),
    "MonthYearSort", YEAR([Date]) * 100 + MONTH([Date])
)
```
Select `MonthYear` → Column tools → **Sort by column** → `MonthYearSort`.
Model view: drag `DateTable[Date]` → `Sales[Order_Date]` (one-to-many). Mark DateTable as the date table.

## 3. Measures (Modeling → New measure)
```DAX
Total Sales   = SUM(Sales[Sales])
Total Profit  = SUM(Sales[Profit])
Profit Margin % = DIVIDE([Total Profit], [Total Sales])
Total Orders  = DISTINCTCOUNT(Sales[Order_ID])
Units Sold    = SUM(Sales[Quantity])
Avg Order Value = DIVIDE([Total Sales], [Total Orders])
Sales LM = CALCULATE([Total Sales], DATEADD(DateTable[Date], -1, MONTH))
MoM Growth % = DIVIDE([Total Sales] - [Sales LM], [Sales LM])
Sales % of Total = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales)))
```

## 4. Dashboard layout (1 page)
| Visual | Fields |
|---|---|
| 4–5 **Cards** (top row) | Total Sales, Total Profit, Profit Margin %, Total Orders, Avg Order Value |
| **Line chart** | X: `DateTable[MonthYear]`, Y: Total Sales (add Total Profit as second line) |
| **Clustered bar** – Sales by Region | Axis: `Sales[Region]`, Values: Total Sales |
| **Bar chart** – Top 10 products | Axis: `Sales[Product]`, Values: Total Sales → Filters pane → Top N = 10 by Total Sales |
| **Donut** – Sales by Category | Legend: `Sales[Category]`, Values: Total Sales |
| **Table / matrix** (optional) | Product, Total Sales, Units Sold, Profit Margin % |
| **Slicers** (left side) | `Region`, `Category`, `Customer_Segment`, `DateTable[Year]` |

Tips: dark-blue title bar, consistent colours, title "Sales Performance Dashboard", and turn on **Edit interactions** so slicers filter every visual.

## 5. Save & publish to GitHub
1. Save as `powerbi/Sales_Dashboard.pbix`.
2. **File → Export → Export to PDF** (optional) and take a screenshot → `powerbi/dashboard_screenshot.png`.
3. Add the screenshot to the top of the README: `![Dashboard](powerbi/dashboard_screenshot.png)`.
