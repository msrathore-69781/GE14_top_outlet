import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 1. Load inputs ────────────────────────────────────────────────────────────
retailer_summary = pd.read_excel("Retailer_Summary.xlsx")
current_month    = pd.read_excel("current_month.xlsx")

# Column aliases (0-indexed)
COL_RETAILER_CODE = current_month.columns[4]   # E — Retailer Code
COL_MATERIAL_CODE = current_month.columns[6]   # G — Material Code (SKU)
COL_SALES_QTY_CM  = current_month.columns[8]   # I — Sales Qty CMCY
COL_SALES_QTY_LY  = current_month.columns[9]   # J — Sales Qty CMLY
COL_TO_ACHIEVED   = current_month.columns[11]  # L — Sales Value CMCY (T.O. Achieved)
COL_TO_BASE       = current_month.columns[12]  # M — Sales Value CMLY (T.O. Base)

# ── 2. Compute per-retailer metrics from current month ────────────────────────
grp = current_month.groupby(COL_RETAILER_CODE)

metrics = pd.DataFrame()

metrics["SKU Base"] = grp.apply(
    lambda x: x.loc[x[COL_SALES_QTY_LY].notna(), COL_MATERIAL_CODE].nunique()
)
metrics["Distt. SKU CM"] = grp.apply(
    lambda x: x.loc[x[COL_SALES_QTY_CM].notna(), COL_MATERIAL_CODE].nunique()
)
metrics["Distt. SKU CM >6EA"] = grp.apply(
    lambda x: x.loc[x[COL_SALES_QTY_CM] > 3, COL_MATERIAL_CODE].nunique()
)
metrics["TO Base"]       = grp[COL_TO_BASE].sum()
metrics["T.O. Achieved"] = grp[COL_TO_ACHIEVED].sum()

metrics = metrics.reset_index()
metrics = metrics.rename(columns={COL_RETAILER_CODE: "Retailer Code"})

# ── 3. Merge with Retailer Summary ───────────────────────────────────────────
summary_cols = retailer_summary[[
    "Retailer Code", "wd_code", "wd", "ret_name",
    "ffr", "Average SKU Count", "Average TO"
]]
df = metrics.merge(summary_cols, on="Retailer Code", how="left")

# ── 4. Derived columns ────────────────────────────────────────────────────────
df["SKU Remaining Target"] = (20 + df["Average SKU Count"]) - df["Distt. SKU CM >6EA"]
df["TO Remaining Target"]  = (1.2 * df["Average TO"]) - df["T.O. Achieved"]

# % TO Achievement = T.O. Achieved / Average TO * 100
df["% TO Achievement"] = df.apply(
    lambda r: (r["T.O. Achieved"] / r["Average TO"] * 100) if r["Average TO"] != 0 else 0,
    axis=1
)

def range_selling_reward(row):
    if row["Distt. SKU CM >6EA"] <= 3:
        return 0
    p = row["Distt. SKU CM >6EA"] - row["Average SKU Count"]
    if p >= 20:   return 2000
    elif p >= 15: return 1500
    elif p >= 10: return 1000
    return 0

def to_reward(row):
    avg_to = row["Average TO"]
    if pd.isna(avg_to) or avg_to == 0:
        return 0
    q = row["T.O. Achieved"] / avg_to - 1
    if q >= 0.20:   return 0.01   * row["T.O. Achieved"]
    elif q >= 0.15: return 0.0075 * row["T.O. Achieved"]
    elif q > 0.10:  return 0.005  * row["T.O. Achieved"]
    return 0

def wdsm_claim(row):
    rs = row["Range Selling Reward"]
    to = row["T.O. Reward"]
    if rs != 0 and to != 0: return 500
    elif rs != 0 or to != 0: return 250
    return 0

df["Range Selling Reward"] = df.apply(range_selling_reward, axis=1)
df["T.O. Reward"]          = df.apply(to_reward, axis=1)
df["WDSM Claim"]           = df.apply(wdsm_claim, axis=1)

# ── 5. Final column order ─────────────────────────────────────────────────────
# [0]  WD Code          [1]  WD Name         [2]  Retailer Code   [3]  Retailer Name
# [4]  FFR              [5]  SKU Base         [6]  Distt. SKU CM   [7]  Distt. SKU CM >6EA
# [8]  Avg SKU Count    [9]  SKU Remaining    [10] TO Base         [11] T.O. Achieved
# [12] Avg TO           [13] % TO Achievement [14] TO Remaining   [15] Range Selling Reward
# [16] T.O. Reward      [17] WDSM Claim

final_cols = [
    "wd_code", "wd", "Retailer Code", "ret_name", "ffr",
    "SKU Base", "Distt. SKU CM", "Distt. SKU CM >6EA",
    "Average SKU Count", "SKU Remaining Target",
    "TO Base", "T.O. Achieved", "Average TO", "% TO Achievement",
    "TO Remaining Target",
    "Range Selling Reward", "T.O. Reward", "WDSM Claim",
]
df_out = df[final_cols].copy()

# ── 6. Write to Excel ─────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Dashboard Backend"

HEADERS = [
    "WD Code", "WD Name", "Retailer Code", "Retailer Name", "FFR",
    "SKU Base", "Distt. SKU CM", "Distt. SKU CM >6EA",
    "Avg SKU Count", "SKU Remaining Target",
    "TO Base", "T.O. Achieved", "Avg TO", "% TO Achievement",
    "TO Remaining Target",
    "Range Selling Reward", "T.O. Reward", "WDSM Claim",
]

hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
hdr_fill  = PatternFill("solid", start_color="1F4E79")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin      = Side(style="thin", color="CCCCCC")
border    = Border(left=thin, right=thin, top=thin, bottom=thin)

fill_info   = PatternFill("solid", start_color="EBF3FB")
fill_sku    = PatternFill("solid", start_color="E2EFDA")
fill_to     = PatternFill("solid", start_color="FFF2CC")
fill_reward = PatternFill("solid", start_color="FCE4D6")

section_fills = {
    0:  fill_info,   1:  fill_info,   2:  fill_info,   3:  fill_info,  4: fill_info,
    5:  fill_sku,    6:  fill_sku,    7:  fill_sku,     8:  fill_sku,   9: fill_sku,
    10: fill_to,     11: fill_to,     12: fill_to,      13: fill_to,   14: fill_to,
    15: fill_reward, 16: fill_reward, 17: fill_reward,
}

for ci, h in enumerate(HEADERS, 1):
    cell = ws.cell(row=1, column=ci, value=h)
    cell.font = hdr_font; cell.fill = hdr_fill
    cell.alignment = hdr_align; cell.border = border

num_font = Font(name="Arial", size=10)
for ri, row in enumerate(df_out.itertuples(index=False), 2):
    for ci, val in enumerate(row, 1):
        cell = ws.cell(row=ri, column=ci, value=val)
        cell.font      = num_font
        cell.fill      = section_fills.get(ci - 1, fill_info)
        cell.border    = border
        cell.alignment = Alignment(
            horizontal="center" if ci > 5 else "left",
            vertical="center"
        )
        # Number formats
        if ci in (11, 12, 13, 15, 16, 17):  # TO + Reward columns
            cell.number_format = '#,##0.00'
        elif ci == 14:                        # % TO Achievement
            cell.number_format = '0.00"%"'

col_widths = [12, 22, 18, 28, 24, 11, 14, 20, 15, 20, 14, 15, 12, 18, 20, 22, 12, 14]
for ci, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(ci)].width = w

ws.row_dimensions[1].height = 40
ws.freeze_panes = "A2"

wb.save("Dashboard_Backend.xlsx")
print("✅ Saved: Dashboard_Backend.xlsx")