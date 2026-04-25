"""
VIỆC 7: Định dạng chuyên nghiệp cho toàn bộ workbook
"""
import win32com.client
import os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

# ─── Color helpers ───────────────────────────────────────
def xlc(hex_color):
    h = hex_color.lstrip('#')
    r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return r + g*256 + b*65536

# Color palette
C_INPUT_BG   = xlc("#D6E4F0")
C_INPUT_FG   = xlc("#1A3A5C")
C_FORM_BG    = xlc("#F5F5F5")
C_FORM_FG    = xlc("#666666")
C_HDR_BG     = xlc("#1A3A5C")
C_HDR_FG     = xlc("#FFFFFF")
C_TITLE_BG   = xlc("#0D2137")
C_TITLE_FG   = xlc("#FFFFFF")
C_ALT_ROW    = xlc("#F9FAFB")
C_POS_BG     = xlc("#E8F8F0")
C_POS_FG     = xlc("#1A7A4A")
C_NEG_BG     = xlc("#FDECEA")
C_NEG_FG     = xlc("#B71C1C")
C_WARN_BG    = xlc("#FFF3E0")
C_WARN_FG    = xlc("#E65100")
C_DANG_BG    = xlc("#FFCDD2")
C_DANG_FG    = xlc("#7F0000")

def style_cell(cell, bg=None, fg=None, bold=False, size=11, italic=False, halign=None):
    if bg is not None: cell.Interior.Color = bg
    if fg is not None: cell.Font.Color = fg
    cell.Font.Bold = bold
    cell.Font.Size = size
    cell.Font.Name = "Calibri"
    if italic: cell.Font.Italic = True
    if halign: cell.HorizontalAlignment = halign

def style_range(rng, bg=None, fg=None, bold=False, size=11):
    if bg is not None: rng.Interior.Color = bg
    if fg is not None: rng.Font.Color = fg
    rng.Font.Bold = bold
    rng.Font.Size = size
    rng.Font.Name = "Calibri"

def add_cf_pos_neg(ws, range_addr):
    """Add conditional formatting: positive=green, negative=red"""
    rng = ws.Range(range_addr)
    rng.FormatConditions.Delete()
    # Positive
    fc = rng.FormatConditions.Add(Type=1, Operator=5, Formula1="0")  # xlCellValue > 0
    fc.Interior.Color = C_POS_BG
    fc.Font.Color = C_POS_FG
    # Negative
    fc2 = rng.FormatConditions.Add(Type=1, Operator=6, Formula1="0")  # xlCellValue < 0
    fc2.Interior.Color = C_NEG_BG
    fc2.Font.Color = C_NEG_FG

# ─────────────────────────────────────────────────────────
print("Opening workbook...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath(file_path), UpdateLinks=False)

# ═══════════════════════════════════════════════════════════
# TAB COLORS
# ═══════════════════════════════════════════════════════════
print("\n[TAB COLORS]")
tab_colors = {
    "Instructions":          xlc("#95A5A6"),  # grey
    "Summary":               xlc("#1A3A5C"),  # dark blue
    "Daily Log":             xlc("#2E86AB"),  # light blue
    "Raw daily data":        xlc("#7B68EE"),  # purple
    "Advanced Setup Planner":xlc("#E67E22"),  # orange
}
for name, color in tab_colors.items():
    try:
        wb.Sheets(name).Tab.Color = color
        print(f"  {name}: tab color set")
    except: print(f"  WARNING: sheet '{name}' not found")

# ═══════════════════════════════════════════════════════════
# SUMMARY SHEET
# ═══════════════════════════════════════════════════════════
print("\n[SUMMARY]")
ws = wb.Sheets("Summary")
pass  # gridlines handled per-window below

# Title row 1 - wrap merge in try/except (may already be merged)
title_rng = ws.Range("A1:K1")
try:
    title_rng.UnMerge()
except: pass
try:
    title_rng.Merge()
except: pass
title_rng.Value = "TRADING PERFORMANCE SUMMARY"
style_range(title_rng, bg=C_TITLE_BG, fg=C_TITLE_FG, bold=True, size=14)
title_rng.HorizontalAlignment = -4108  # xlCenter
ws.Rows(1).RowHeight = 36

# Input cells (blue)
for addr in ["B2","B3","B5","B8","B9","B10"]:
    style_cell(ws.Range(addr), bg=C_INPUT_BG, fg=C_INPUT_FG)

# Formula cells (grey) - rows 4,6,7,11-19
formula_rows = [4,6,7,11,12,13,14,15,16,17,18,19]
for r in formula_rows:
    style_cell(ws.Cells(r,2), bg=C_FORM_BG, fg=C_FORM_FG)

# Label column A
style_range(ws.Range("A2:A19"), fg=C_INPUT_FG)

# Weekly table header row 21
style_range(ws.Range("A21:K21"), bg=C_HDR_BG, fg=C_HDR_FG, bold=True, size=12)
ws.Rows(21).RowHeight = 24

# Weekly table alternating rows 22-25
for r in range(22, 26):
    bg = C_ALT_ROW if r % 2 == 0 else xlc("#FFFFFF")
    style_range(ws.Range(f"A{r}:K{r}"), bg=bg)

# Conditional formatting: weekly P&L (col F and H = cols 6,8)
add_cf_pos_neg(ws, "F22:F25")
add_cf_pos_neg(ws, "H22:H25")
add_cf_pos_neg(ws, "B11")  # Cumulative P&L

# Number formats
ws.Range("B2").NumberFormat = "$#,##0.00"
ws.Range("B10").NumberFormat = "$#,##0.00"
ws.Range("B11").NumberFormat = "$#,##0.00"
ws.Range("B12").NumberFormat = "$#,##0.00"
ws.Range("B8").NumberFormat = "0.0%"

# Freeze row 1
ws.Activate()
excel.ActiveWindow.FreezePanes = False
ws.Range("A2").Select()
excel.ActiveWindow.FreezePanes = True
print("  Summary: done")

# ═══════════════════════════════════════════════════════════
# DAILY LOG
# ═══════════════════════════════════════════════════════════
print("\n[DAILY LOG]")
ws = wb.Sheets("Daily Log")
pass  # gridlines handled per-window below

# Header row 1
style_range(ws.Range("A1:Y1"), bg=C_HDR_BG, fg=C_HDR_FG, bold=True, size=12)
ws.Rows(1).RowHeight = 24

# Input cols A,B,D,M (1,2,4,13)
for r in range(2, 101):
    for col in [1, 2, 4, 13]:
        ws.Cells(r, col).Interior.Color = C_INPUT_BG
        ws.Cells(r, col).Font.Color = C_INPUT_FG

# Formula cols C,E-L,N-Y
formula_cols = [3] + list(range(5,13)) + list(range(14,26))
for r in range(2, 101):
    for col in formula_cols:
        ws.Cells(r, col).Interior.Color = C_FORM_BG
        ws.Cells(r, col).Font.Color = C_FORM_FG

# Conditional: col H (Net P&L = col 8)
add_cf_pos_neg(ws, "H2:H100")

# Conditional: col V (Losses vs Cap = col 22) - warning/danger
rng_v = ws.Range("V2:V100")
rng_v.FormatConditions.Delete()
fc_warn = rng_v.FormatConditions.Add(Type=1, Operator=5, Formula1="0.8")
fc_warn.Interior.Color = C_WARN_BG; fc_warn.Font.Color = C_WARN_FG
fc_dang = rng_v.FormatConditions.Add(Type=1, Operator=6, Formula1="1.001")
fc_dang.Interior.Color = C_DANG_BG; fc_dang.Font.Color = C_DANG_FG

# Column widths
col_widths = {1:12, 2:8, 3:12, 4:8, 13:25}
for i in range(5, 13): col_widths[i] = 11
for i in range(14, 26): col_widths[i] = 9
for col, w in col_widths.items():
    ws.Columns(col).ColumnWidth = w

# Freeze row 1
ws.Activate()
excel.ActiveWindow.FreezePanes = False
ws.Range("A2").Select()
excel.ActiveWindow.FreezePanes = True
print("  Daily Log: done")

# ═══════════════════════════════════════════════════════════
# RAW DAILY DATA
# ═══════════════════════════════════════════════════════════
print("\n[RAW DAILY DATA]")
ws = wb.Sheets("Raw daily data")
pass  # gridlines handled per-window below

# Header row 1
style_range(ws.Range("A1:J1"), bg=C_HDR_BG, fg=C_HDR_FG, bold=True, size=12)
ws.Rows(1).RowHeight = 24

# Add paste instruction to header
ws.Range("A1").Value = "Symbol  ← Paste broker data into columns A–H only. Do not edit I–J →"

# Input cols A-H (1-8)
for r in range(2, 501):
    for col in range(1, 9):
        ws.Cells(r, col).Interior.Color = C_INPUT_BG

# Formula cols I-J (9-10)
for r in range(2, 501):
    for col in [9, 10]:
        ws.Cells(r, col).Interior.Color = C_FORM_BG
        ws.Cells(r, col).Font.Color = C_FORM_FG

# Freeze
ws.Activate()
excel.ActiveWindow.FreezePanes = False
ws.Range("A2").Select()
excel.ActiveWindow.FreezePanes = True
print("  Raw daily data: done")

# ═══════════════════════════════════════════════════════════
# ADVANCED SETUP PLANNER
# ═══════════════════════════════════════════════════════════
print("\n[ADVANCED SETUP PLANNER]")
ws = wb.Sheets("Advanced Setup Planner")

# Title row 1 (if exists)
try:
    ws.Range("A1:M1").Merge()
    ws.Range("A1").Value = "ADVANCED SETUP PLANNER"
    style_range(ws.Range("A1:M1"), bg=C_TITLE_BG, fg=C_TITLE_FG, bold=True, size=14)
    ws.Rows(1).RowHeight = 36
except: pass

# Input cells (col B rows 5-21, col E rows 5-18)
for r in range(5, 22):
    ws.Cells(r, 2).Interior.Color = C_INPUT_BG
    ws.Cells(r, 2).Font.Color = C_INPUT_FG
for r in range(5, 19):
    ws.Cells(r, 5).Interior.Color = C_INPUT_BG
    ws.Cells(r, 5).Font.Color = C_INPUT_FG

# BOS/CHOCH inputs
for r in range(24, 28):
    ws.Cells(r, 3).Interior.Color = C_INPUT_BG
    ws.Cells(r, 3).Font.Color = C_INPUT_FG
    ws.Cells(r, 4).Interior.Color = C_INPUT_BG
    ws.Cells(r, 4).Font.Color = C_INPUT_FG

# Conditional: Side (B31) - LONG=green, SHORT=red
rng_side = ws.Range("B31")
rng_side.FormatConditions.Delete()
fc_long = rng_side.FormatConditions.Add(Type=1, Operator=3, Formula1='"LONG"')
fc_long.Interior.Color = C_POS_BG; fc_long.Font.Color = C_POS_FG; fc_long.Font.Bold = True
fc_short = rng_side.FormatConditions.Add(Type=1, Operator=3, Formula1='"SHORT"')
fc_short.Interior.Color = C_NEG_BG; fc_short.Font.Color = C_NEG_FG; fc_short.Font.Bold = True

# Conditional: Trade Mode (B33)
rng_mode = ws.Range("B33")
rng_mode.FormatConditions.Delete()
for val, bg, fg in [
    ('"CoreLong"', C_POS_BG, C_POS_FG),
    ('"A+Long"', C_POS_BG, C_POS_FG),
    ('"CTshort"', C_NEG_BG, C_NEG_FG),
    ('"Avoid"', C_DANG_BG, C_DANG_FG),
]:
    fc = rng_mode.FormatConditions.Add(Type=1, Operator=3, Formula1=val)
    fc.Interior.Color = bg; fc.Font.Color = fg; fc.Font.Bold = True

# Group engine columns M-AK (13-37)
try:
    ws.Columns("M:AK").Group()
    print("  Engine columns M:AK grouped")
except Exception as e:
    print(f"  Group warning: {e}")

# Freeze row 1
ws.Activate()
excel.ActiveWindow.FreezePanes = False
ws.Range("A2").Select()
excel.ActiveWindow.FreezePanes = True
print("  Advanced Setup Planner: done")

# ═══════════════════════════════════════════════════════════
# REMOVE GRIDLINES ON ALL SHEETS via ActiveWindow
# ═══════════════════════════════════════════════════════════
print("\n[GRIDLINES]")
excel.Visible = True  # need visible window to set gridlines
sheet_names = ["Instructions","Summary","Daily Log","Raw daily data","Advanced Setup Planner"]
for sname in sheet_names:
    try:
        wb.Sheets(sname).Activate()
        excel.ActiveWindow.DisplayGridlines = False
        print(f"  {sname}: gridlines off")
    except Exception as e:
        print(f"  WARNING {sname}: {e}")
excel.Visible = False

# ─────────────────────────────────────────────────────────
print("\nSaving...")
wb.Save()
wb.Close()
excel.Quit()
print("\n=== VIỆC 7 HOÀN THÀNH ===")
