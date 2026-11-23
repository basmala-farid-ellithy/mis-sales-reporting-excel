# MIS Sales Reporting (Excel)

A reproducible build of an MIS-style monthly sales report: the recurring Excel pack an
MIS analyst produces for management - KPI summary, monthly trend, product-line
performance, branch share and payment mix, with native Excel charts.

## What it builds
`build_report.py` reads the raw sales CSV and produces `MIS_Sales_Report_Q1_2019.xlsx`:

| Sheet | Content |
|---|---|
| KPI Summary | Total revenue, gross income, transactions, average basket, average rating |
| Monthly Trend | Month-over-month revenue + bar chart |
| Product Lines | Revenue and gross income per product line + chart |
| Branch and Payment | Branch revenue share (pie) and payment mix by customer type |

## Headline numbers (Q1 2019)
- Total revenue: 322,966.75 across 1,000 transactions (avg basket 322.97)
- January strongest (116.3k), February dip (97.2k), March rebound (109.5k)
- Revenue evenly split across the three branches; Food and beverages is the top line

## Run it
```
pip install pandas openpyxl
python build_report.py
```

## Data
Public "Supermarket Sales" dataset (plotly/datasets mirror of the well-known Kaggle dataset),
included in this repo for reproducibility.

## Stack
Python (pandas, openpyxl) generating a native Excel workbook - mirroring an MIS reporting
workflow that can equally be rebuilt in Power BI from the same aggregations.
