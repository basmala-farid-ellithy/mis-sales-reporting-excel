"""Build an MIS-style monthly sales report workbook (openpyxl): a KPI summary sheet,
monthly and product-line pivot sheets, and native Excel charts - the kind of recurring
report produced in an MIS role."""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_csv("supermarket_sales.csv", parse_dates=["Date"])
df["Month"] = df["Date"].dt.strftime("%Y-%m")

wb = Workbook()
hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(color="FFFFFF", bold=True)

def style_header(ws, row=1):
    for c in ws[row]:
        c.fill = hdr_fill; c.font = hdr_font

def add_df(ws, d):
    for r in dataframe_to_rows(d, index=False):
        ws.append(r)
    style_header(ws)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value or "")) for c in col) + 3

# Sheet 1: KPI summary
ws = wb.active; ws.title = "KPI Summary"
kpis = pd.DataFrame({
    "KPI": ["Total revenue", "Gross income", "Transactions", "Avg basket value", "Avg rating"],
    "Value": [round(df["Total"].sum(), 2), round(df["Gross income"].sum(), 2), len(df),
              round(df["Total"].mean(), 2), round(df["Customer stratification rating"].mean(), 2)],
})
add_df(ws, kpis)

# Sheet 2: Monthly trend + bar chart
ws2 = wb.create_sheet("Monthly Trend")
monthly = df.groupby("Month", as_index=False)["Total"].sum().round(2)
monthly.columns = ["Month", "Revenue"]
add_df(ws2, monthly)
ch = BarChart(); ch.title = "Monthly Revenue"; ch.y_axis.title = "Revenue"
ch.add_data(Reference(ws2, min_col=2, min_row=1, max_row=1 + len(monthly)), titles_from_data=True)
ch.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=1 + len(monthly)))
ws2.add_chart(ch, "D2")

# Sheet 3: Product line performance + chart
ws3 = wb.create_sheet("Product Lines")
pl = df.groupby("Product line", as_index=False)[["Total", "Gross income"]].sum().round(2).sort_values("Total", ascending=False)
pl.columns = ["Product line", "Revenue", "Gross income"]
add_df(ws3, pl)
ch2 = BarChart(); ch2.type = "col"; ch2.title = "Revenue by Product Line"
ch2.add_data(Reference(ws3, min_col=2, min_row=1, max_row=1 + len(pl)), titles_from_data=True)
ch2.set_categories(Reference(ws3, min_col=1, min_row=2, max_row=1 + len(pl)))
ws3.add_chart(ch2, "E2")

# Sheet 4: Branch & payment mix + pie
ws4 = wb.create_sheet("Branch and Payment")
br = df.groupby("City", as_index=False)["Total"].sum().round(2)
br.columns = ["City", "Revenue"]
add_df(ws4, br)
pie = PieChart(); pie.title = "Revenue Share by Branch"
pie.add_data(Reference(ws4, min_col=2, min_row=1, max_row=1 + len(br)), titles_from_data=True)
pie.set_categories(Reference(ws4, min_col=1, min_row=2, max_row=1 + len(br)))
ws4.add_chart(pie, "D2")
row0 = len(br) + 4
pm = df.pivot_table(index="Payment", columns="Customer type", values="Total", aggfunc="sum").round(2).reset_index()
ws4.cell(row=row0 - 1, column=1, value="Payment mix by customer type").font = Font(bold=True)
for r in dataframe_to_rows(pm, index=False):
    ws4.append(r)
style_header(ws4, row=row0)

wb.save("MIS_Sales_Report_Q1_2019.xlsx")
print("workbook saved")
