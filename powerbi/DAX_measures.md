# Power BI Dashboard - Measures Used

Data source: `data/cleaned_sales_data.csv` (table: `cleaned_sales_data`)

## Measures
Total Sales = SUM(cleaned_sales_data[Sales])
Total Profit = SUM(cleaned_sales_data[Profit])
Profit Margin % = DIVIDE([Total Profit], [Total Sales])
Total Orders = DISTINCTCOUNT(cleaned_sales_data[Order_ID])

## Visuals
- KPI cards: Total Sales, Total Profit, Profit Margin %, Total Orders
- Line chart: Total Sales by Month
- Bar chart: Total Sales by Region
- Bar chart: Total Sales by Product
- Slicers: Region, Category

Tips: dark-blue title bar, consistent colours, title "Sales Performance Dashboard", and turn on **Edit interactions** so slicers filter every visual.

## 5. Save & publish to GitHub
1. Save as `powerbi/Sales_Dashboard.pbix`.
2. **File → Export → Export to PDF** (optional) and take a screenshot → `powerbi/dashboard_screenshot.png`.
3. Add the screenshot to the top of the README: `![Dashboard](powerbi/dashboard_screenshot.png)`.
