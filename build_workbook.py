import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

# ---------- Load raw data ----------
with open("/home/claude/sales_dashboard/sales_raw.csv") as f:
    reader = csv.DictReader(f)
    data = list(reader)

n = len(data)  # 900
last_row = n + 1  # header is row 1, data rows 2..last_row

regions = ["North", "South", "East", "West"]
categories = ["Electronics", "Clothing", "Home & Kitchen", "Sports"]
products = sorted(set(r["Product"] for r in data))
product_category = {}
for r in data:
    product_category[r["Product"]] = r["Category"]

# ---------- Styles ----------
NAVY = "1F3864"
BLUE = "2E5EAA"
LIGHT_BLUE_FILL = PatternFill("solid", fgColor="DCE6F1")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
KPI_FILL = PatternFill("solid", fgColor="2E5EAA")
FONT_NAME = "Arial"

title_font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
header_font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
label_font = Font(name=FONT_NAME, size=10, bold=True, color="404040")
kpi_value_font = Font(name=FONT_NAME, size=20, bold=True, color="FFFFFF")
kpi_label_font = Font(name=FONT_NAME, size=10, bold=False, color="DCE6F1")
normal_font = Font(name=FONT_NAME, size=10)
section_font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
thin = Side(style="thin", color="B7B7B7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

# =====================================================================
# SHEET 1: Raw Data
# =====================================================================
ws_raw = wb.active
ws_raw.title = "Raw Data"

headers = ["Date", "Region", "Category", "Product", "Units", "Unit Price", "Revenue"]
for col, h in enumerate(headers, start=1):
    c = ws_raw.cell(row=1, column=col, value=h)
    c.font = header_font
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")

for i, r in enumerate(data, start=2):
    ws_raw.cell(row=i, column=1, value=r["Date"])
    ws_raw.cell(row=i, column=2, value=r["Region"])
    ws_raw.cell(row=i, column=3, value=r["Category"])
    ws_raw.cell(row=i, column=4, value=r["Product"])
    ws_raw.cell(row=i, column=5, value=int(r["Units"]))
    ws_raw.cell(row=i, column=6, value=float(r["UnitPrice"]))
    ws_raw.cell(row=i, column=7, value=f"=E{i}*F{i}")
    ws_raw.cell(row=i, column=1).number_format = "yyyy-mm-dd"
    ws_raw.cell(row=i, column=6).number_format = "#,##0.00"
    ws_raw.cell(row=i, column=7).number_format = "#,##0.00"

widths = [12, 10, 16, 24, 8, 12, 14]
for col, w in enumerate(widths, start=1):
    ws_raw.column_dimensions[get_column_letter(col)].width = w
ws_raw.freeze_panes = "A2"

RAW = "'Raw Data'"
DATE_RNG = f"{RAW}!$A$2:$A${last_row}"
REGION_RNG = f"{RAW}!$B$2:$B${last_row}"
CATEGORY_RNG = f"{RAW}!$C$2:$C${last_row}"
PRODUCT_RNG = f"{RAW}!$D$2:$D${last_row}"
UNITS_RNG = f"{RAW}!$E$2:$E${last_row}"
REVENUE_RNG = f"{RAW}!$G$2:$G${last_row}"

# =====================================================================
# SHEET 2: Product Summary (helper sheet for Top-Selling Products)
# =====================================================================
ws_ps = wb.create_sheet("Product Summary")
ps_headers = ["Product", "Category", "Units Sold", "Revenue", "Rank"]
for col, h in enumerate(ps_headers, start=1):
    c = ws_ps.cell(row=1, column=col, value=h)
    c.font = header_font
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")

p_last = len(products) + 1
for i, p in enumerate(products, start=2):
    ws_ps.cell(row=i, column=1, value=p)
    ws_ps.cell(row=i, column=2, value=product_category[p])
    ws_ps.cell(row=i, column=3, value=f'=SUMIF({PRODUCT_RNG},A{i},{UNITS_RNG})')
    ws_ps.cell(row=i, column=4, value=f'=SUMIF({PRODUCT_RNG},A{i},{REVENUE_RNG})')
    ws_ps.cell(row=i, column=4).number_format = "#,##0.00"
    ws_ps.cell(row=i, column=5,
               value=f'=RANK(D{i},$D$2:$D${p_last})+COUNTIF($D$2:D{i},D{i})-1')

widths_ps = [26, 16, 12, 14, 8]
for col, w in enumerate(widths_ps, start=1):
    ws_ps.column_dimensions[get_column_letter(col)].width = w
ws_ps.freeze_panes = "A2"

PS = "'Product Summary'"
PS_PRODUCT_RNG = f"{PS}!$A$2:$A${p_last}"
PS_REVENUE_RNG = f"{PS}!$D$2:$D${p_last}"
PS_UNITS_RNG = f"{PS}!$C$2:$C${p_last}"
PS_RANK_RNG = f"{PS}!$E$2:$E${p_last}"

# =====================================================================
# SHEET 3: Dashboard
# =====================================================================
ws = wb.create_sheet("Dashboard")
ws.sheet_view.showGridLines = False

# --- Title band ---
ws.merge_cells("A1:L2")
ws["A1"] = "SALES PERFORMANCE DASHBOARD — FY2024"
ws["A1"].font = title_font
ws["A1"].fill = HEADER_FILL
ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 22
ws.row_dimensions[2].height = 22

# --- Filters ---
ws["A4"] = "Region Filter:"
ws["A4"].font = label_font
ws["B4"] = "All"
ws["B4"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F3864")
ws["B4"].fill = PatternFill("solid", fgColor="FFF2CC")
ws["B4"].border = border

ws["D4"] = "Category Filter:"
ws["D4"].font = label_font
ws["E4"] = "All"
ws["E4"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F3864")
ws["E4"].fill = PatternFill("solid", fgColor="FFF2CC")
ws["E4"].border = border

dv_region = DataValidation(type="list", formula1='"All,North,South,East,West"', allow_blank=False)
dv_category = DataValidation(type="list",
                              formula1='"All,Electronics,Clothing,Home & Kitchen,Sports"',
                              allow_blank=False)
ws.add_data_validation(dv_region)
ws.add_data_validation(dv_category)
dv_region.add(ws["B4"])
dv_category.add(ws["E4"])

ws["G4"] = "(Choose 'All' or a specific value — every KPI, the monthly trend, and top products update automatically.)"
ws["G4"].font = Font(name=FONT_NAME, size=9, italic=True, color="808080")

REG_FILTER = "$B$4"
CAT_FILTER = "$E$4"

def filt_mask():
    return (f'({REGION_RNG}=IF({REG_FILTER}="All",{REGION_RNG},{REG_FILTER}))*'
            f'({CATEGORY_RNG}=IF({CAT_FILTER}="All",{CATEGORY_RNG},{CAT_FILTER}))')

mask = filt_mask()

# --- KPI cards ---
kpi_defs = [
    ("TOTAL REVENUE", f'=SUMPRODUCT({mask}*{REVENUE_RNG})', "#,##0.00"),
    ("TOTAL UNITS SOLD", f'=SUMPRODUCT({mask}*{UNITS_RNG})', "#,##0"),
    ("NUMBER OF ORDERS", f'=SUMPRODUCT({mask}*1)', "#,##0"),
    ("AVG ORDER VALUE", None, "#,##0.00"),  # filled after revenue/orders cells known
]

kpi_start_col = 1  # A
kpi_width = 3  # each kpi spans 3 columns
kpi_row_top = 6
kpi_row_bottom = 8

kpi_cells_value = {}
for idx, (label, formula, numfmt) in enumerate(kpi_defs):
    start_col = kpi_start_col + idx * kpi_width
    end_col = start_col + kpi_width - 1
    start_letter = get_column_letter(start_col)
    end_letter = get_column_letter(end_col)

    ws.merge_cells(f"{start_letter}{kpi_row_top}:{end_letter}{kpi_row_top}")
    lbl_cell = ws[f"{start_letter}{kpi_row_top}"]
    lbl_cell.value = label
    lbl_cell.font = kpi_label_font
    lbl_cell.fill = KPI_FILL
    lbl_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"{start_letter}{kpi_row_top+1}:{end_letter}{kpi_row_bottom}")
    val_cell = ws[f"{start_letter}{kpi_row_top+1}"]
    val_cell.font = kpi_value_font
    val_cell.fill = KPI_FILL
    val_cell.alignment = Alignment(horizontal="center", vertical="center")
    val_cell.number_format = numfmt
    kpi_cells_value[label] = f"{start_letter}{kpi_row_top+1}"
    if formula:
        val_cell.value = formula

    for r in (kpi_row_top, kpi_row_top+1, kpi_row_top+2):
        for c in range(start_col, end_col+1):
            ws.cell(row=r, column=c).fill = KPI_FILL

ws.row_dimensions[kpi_row_top].height = 16
ws.row_dimensions[kpi_row_top+1].height = 20
ws.row_dimensions[kpi_row_top+2].height = 20

# Avg order value = total revenue / number of orders
rev_cell = kpi_cells_value["TOTAL REVENUE"]
ord_cell = kpi_cells_value["NUMBER OF ORDERS"]
avg_cell_ref = kpi_cells_value["AVG ORDER VALUE"]
ws[avg_cell_ref] = f'=IF({ord_cell}=0,0,{rev_cell}/{ord_cell})'

# =====================================================================
# Monthly Revenue Trend (respects filters)
# =====================================================================
month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

mt_row_header = 11
ws.cell(row=mt_row_header - 1, column=1, value="MONTHLY REVENUE TREND").font = section_font

ws.cell(row=mt_row_header, column=1, value="Month").font = header_font
ws.cell(row=mt_row_header, column=1).fill = HEADER_FILL
ws.cell(row=mt_row_header, column=2, value="Revenue").font = header_font
ws.cell(row=mt_row_header, column=2).fill = HEADER_FILL

for i, m in enumerate(month_names, start=1):
    r = mt_row_header + i
    ws.cell(row=r, column=1, value=m)
    formula = (f'=SUMPRODUCT((MONTH({DATE_RNG})={i})*{mask}*{REVENUE_RNG})')
    c = ws.cell(row=r, column=2, value=formula)
    c.number_format = "#,##0"
    ws.cell(row=r, column=1).border = border
    c.border = border

mt_row_last = mt_row_header + 12

# =====================================================================
# Revenue by Region (overall, not filtered — for comparison)
# =====================================================================
reg_row_header = mt_row_header
reg_col = 4  # column D
ws.cell(row=reg_row_header - 1, column=reg_col, value="REVENUE BY REGION").font = section_font
ws.cell(row=reg_row_header, column=reg_col, value="Region").font = header_font
ws.cell(row=reg_row_header, column=reg_col).fill = HEADER_FILL
ws.cell(row=reg_row_header, column=reg_col+1, value="Revenue").font = header_font
ws.cell(row=reg_row_header, column=reg_col+1).fill = HEADER_FILL

for i, reg in enumerate(regions, start=1):
    r = reg_row_header + i
    ws.cell(row=r, column=reg_col, value=reg).border = border
    c = ws.cell(row=r, column=reg_col+1, value=f'=SUMIF({REGION_RNG},{get_column_letter(reg_col)}{r},{REVENUE_RNG})')
    c.number_format = "#,##0"
    c.border = border

reg_row_last = reg_row_header + len(regions)

# =====================================================================
# Revenue by Category (overall, not filtered)
# =====================================================================
cat_col = 7  # column G
ws.cell(row=reg_row_header - 1, column=cat_col, value="REVENUE BY CATEGORY").font = section_font
ws.cell(row=reg_row_header, column=cat_col, value="Category").font = header_font
ws.cell(row=reg_row_header, column=cat_col).fill = HEADER_FILL
ws.cell(row=reg_row_header, column=cat_col+1, value="Revenue").font = header_font
ws.cell(row=reg_row_header, column=cat_col+1).fill = HEADER_FILL

for i, cat in enumerate(categories, start=1):
    r = reg_row_header + i
    ws.cell(row=r, column=cat_col, value=cat).border = border
    c = ws.cell(row=r, column=cat_col+1, value=f'=SUMIF({CATEGORY_RNG},{get_column_letter(cat_col)}{r},{REVENUE_RNG})')
    c.number_format = "#,##0"
    c.border = border

cat_row_last = reg_row_header + len(categories)

# =====================================================================
# Top 10 Products by Revenue (overall)
# =====================================================================
top_col = 10  # column J
top_row_header = reg_row_header
ws.cell(row=top_row_header - 1, column=top_col, value="TOP 10 PRODUCTS BY REVENUE").font = section_font
ws.cell(row=top_row_header, column=top_col, value="Product").font = header_font
ws.cell(row=top_row_header, column=top_col).fill = HEADER_FILL
ws.cell(row=top_row_header, column=top_col+1, value="Revenue").font = header_font
ws.cell(row=top_row_header, column=top_col+1).fill = HEADER_FILL

for k in range(1, 11):
    r = top_row_header + k
    ws.cell(row=r, column=top_col,
            value=f'=INDEX({PS_PRODUCT_RNG},MATCH({k},{PS_RANK_RNG},0))').border = border
    c = ws.cell(row=r, column=top_col+1,
                value=f'=INDEX({PS_REVENUE_RNG},MATCH({k},{PS_RANK_RNG},0))')
    c.number_format = "#,##0"
    c.border = border

top_row_last = top_row_header + 10

# --- column widths for dashboard ---
dash_widths = {1: 14, 2: 12, 3: 3, 4: 10, 5: 12, 6: 3, 7: 16, 8: 12, 9: 3, 10: 24, 11: 12, 12: 3}
for col, w in dash_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# =====================================================================
# Charts
# =====================================================================
chart_anchor_row = max(mt_row_last, cat_row_last, top_row_last) + 3

# Line chart: Monthly trend
line = LineChart()
line.title = "Monthly Revenue Trend"
line.style = 2
line.y_axis.title = "Revenue"
line.x_axis.title = "Month"
line.height = 8
line.width = 16
data_ref = Reference(ws, min_col=2, min_row=mt_row_header, max_row=mt_row_last)
cats_ref = Reference(ws, min_col=1, min_row=mt_row_header+1, max_row=mt_row_last)
line.add_data(data_ref, titles_from_data=True)
line.set_categories(cats_ref)
line.legend = None
ws.add_chart(line, f"A{chart_anchor_row}")

# Bar chart: Revenue by Region
bar_region = BarChart()
bar_region.title = "Revenue by Region"
bar_region.style = 10
bar_region.y_axis.title = "Revenue"
bar_region.height = 8
bar_region.width = 12
data_ref = Reference(ws, min_col=reg_col+1, min_row=reg_row_header, max_row=reg_row_last)
cats_ref = Reference(ws, min_col=reg_col, min_row=reg_row_header+1, max_row=reg_row_last)
bar_region.add_data(data_ref, titles_from_data=True)
bar_region.set_categories(cats_ref)
bar_region.legend = None
ws.add_chart(bar_region, f"E{chart_anchor_row}")

# Bar chart: Revenue by Category
bar_cat = BarChart()
bar_cat.title = "Revenue by Category"
bar_cat.style = 11
bar_cat.y_axis.title = "Revenue"
bar_cat.height = 8
bar_cat.width = 14
data_ref = Reference(ws, min_col=cat_col+1, min_row=reg_row_header, max_row=cat_row_last)
cats_ref = Reference(ws, min_col=cat_col, min_row=reg_row_header+1, max_row=cat_row_last)
bar_cat.add_data(data_ref, titles_from_data=True)
bar_cat.set_categories(cats_ref)
bar_cat.legend = None
ws.add_chart(bar_cat, f"I{chart_anchor_row}")

chart_anchor_row2 = chart_anchor_row + 17

# Bar chart: Top 10 products
bar_top = BarChart()
bar_top.type = "bar"
bar_top.title = "Top 10 Products by Revenue"
bar_top.style = 12
bar_top.x_axis.title = "Revenue"
bar_top.height = 9
bar_top.width = 18
data_ref = Reference(ws, min_col=top_col+1, min_row=top_row_header, max_row=top_row_last)
cats_ref = Reference(ws, min_col=top_col, min_row=top_row_header+1, max_row=top_row_last)
bar_top.add_data(data_ref, titles_from_data=True)
bar_top.set_categories(cats_ref)
bar_top.legend = None
ws.add_chart(bar_top, f"A{chart_anchor_row2}")

# =====================================================================
# SHEET 4: README
# =====================================================================
ws_r = wb.create_sheet("README")
ws_r["A1"] = "Sales Dashboard Analysis — Project Notes"
ws_r["A1"].font = Font(name=FONT_NAME, size=14, bold=True, color=NAVY)
notes = [
    "",
    "Dataset: 900 synthetic sales transactions for FY2024, generated for this project",
    "(no real company data used). Columns: Date, Region, Category, Product, Units, Unit Price.",
    "",
    "Sheets:",
    "  Raw Data        - transaction-level data; Revenue = Units x Unit Price (formula).",
    "  Product Summary - helper sheet aggregating Units/Revenue per product, used to rank",
    "                    products for the Top 10 Products chart (RANK + INDEX/MATCH).",
    "  Dashboard        - KPIs, Region/Category filters (data-validation dropdowns), monthly",
    "                    trend, and four charts.",
    "",
    "How the filters work:",
    "  Change the Region or Category dropdown on the Dashboard (cells B4 and E4) to 'All' or",
    "  a specific value. Every KPI and the Monthly Revenue Trend recalculate instantly using",
    "  SUMPRODUCT formulas that check the filter cells - this mimics how a Power BI slicer",
    "  drives visuals, built natively in Excel.",
    "",
    "Concepts demonstrated: data cleaning & transformation, KPI design, formula-driven",
    "'slicers', dashboard layout, and business insight extraction from sales data.",
]
for i, line_txt in enumerate(notes, start=2):
    ws_r.cell(row=i, column=1, value=line_txt).font = normal_font
ws_r.column_dimensions["A"].width = 100

wb.save("/home/claude/sales_dashboard/Sales_Dashboard_Analysis.xlsx")
print("Workbook saved.")
