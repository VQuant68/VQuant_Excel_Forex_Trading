"""
FINAL CLEAN SCRIPT
Base: Trading_Workbook_FINAL_v2.xlsx
Output: Trading_Workbook_FINAL.xlsx

Chỉ làm 4 việc, KHÔNG đụng layout:
  1. Xóa Named Ranges bị hỏng
  2. Thêm Data Validations đúng vị trí
  3. Tạo sheet Instructions
  4. Ẩn Backend (VeryHidden)
"""

import win32com.client
import os
import shutil

# ─────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────
def h2rgb(hex_color):
    h = hex_color.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def rgb2xl(rgb):
    return rgb[0] + (rgb[1] * 256) + (rgb[2] * 65536)

def xlc(hex_color):
    return rgb2xl(h2rgb(hex_color))

def add_dropdown(rng, formula):
    """Add list validation to a range."""
    rng.Validation.Delete()
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula)
    rng.Validation.InCellDropdown = True

def find_cell(ws, search_text, search_col=None):
    """Find a cell containing search_text. Returns cell or None."""
    try:
        found = ws.UsedRange.Find(
            What=search_text,
            LookIn=-4163,  # xlValues
            LookAt=2,      # xlPart
        )
        return found
    except:
        return None

# ─────────────────────────────────────────────────────────
# STEP 1: Setup
# ─────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("FINAL CLEAN SCRIPT — Starting...")
    print("=" * 60)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    src = os.path.abspath("Trading_Workbook_FINAL_v2.xlsx")
    dst = os.path.abspath("Trading_Workbook_MASTER.xlsx")

    if os.path.exists(dst):
        os.remove(dst)
    shutil.copy(src, dst)
    print(f"Copied v2 → {dst}")

    wb = excel.Workbooks.Open(dst)

    # ─────────────────────────────────────────────────────
    # STEP 2: Fix Named Ranges
    # ─────────────────────────────────────────────────────
    print("\n[1] Fixing Named Ranges...")
    broken = ["tblPolicy", "Macro_Data", "tblMacroUS"]
    for name in broken:
        try:
            wb.Names(name).Delete()
            print(f"    Deleted: {name}")
        except:
            pass
    # Verify the 7 LT_ ranges exist
    expected = ["LT_TrendBull","LT_WeekNum","LT_MacroDir",
                "LT_RatePolicy","LT_RiskEnv","LT_CategoryUS","LT_Speakers"]
    ok = 0
    for e in expected:
        try:
            wb.Names(e)
            ok += 1
        except:
            print(f"    MISSING Named Range: {e}")
    print(f"    {ok}/7 LT_ Named Ranges present.")

    # ─────────────────────────────────────────────────────
    # STEP 3: Data Validations
    # ─────────────────────────────────────────────────────
    print("\n[2] Adding Data Validations...")

    # 3a. Daily Log column D (Y/N)
    sh_daily = wb.Sheets("Daily Log")
    add_dropdown(sh_daily.Range("D2:D100"), "Y,N")
    print("    Daily Log D2:D100 → Y,N")

    # 3b. Raw daily data column C (buy/sell)
    sh_raw = wb.Sheets("Raw daily data")
    add_dropdown(sh_raw.Range("C2:C500"), "buy,sell")
    print("    Raw daily data C2:C500 → buy,sell")

    # 3c. Advanced Setup Planner — find Last BOS / Last CHOCH cells
    # In v7, HUD is at column A (moved from AN). BOS/CHOCH headers are in the HUD.
    sh_planner = wb.Sheets("Advanced Setup Planner")
    bos_cell = find_cell(sh_planner, "Last BOS")
    choch_cell = find_cell(sh_planner, "Last CHOCH")

    if bos_cell:
        bos_row = bos_cell.Row
        bos_col = bos_cell.Column
        # Input cells are rows below header (4 rows: 1D/4H/1H/15M)
        input_bos = sh_planner.Range(
            sh_planner.Cells(bos_row + 1, bos_col),
            sh_planner.Cells(bos_row + 4, bos_col)
        )
        add_dropdown(input_bos, "=LT_TrendBull")
        print(f"    Last BOS inputs ({bos_cell.Address}, +4 rows) → =LT_TrendBull")
    else:
        print("    WARNING: Could not find 'Last BOS' header cell")

    if choch_cell:
        choch_row = choch_cell.Row
        choch_col = choch_cell.Column
        input_choch = sh_planner.Range(
            sh_planner.Cells(choch_row + 1, choch_col),
            sh_planner.Cells(choch_row + 4, choch_col)
        )
        add_dropdown(input_choch, "=LT_TrendBull")
        print(f"    Last CHOCH inputs ({choch_cell.Address}, +4 rows) → =LT_TrendBull")
    else:
        print("    WARNING: Could not find 'Last CHOCH' header cell")

    # ─────────────────────────────────────────────────────
    # STEP 4: Create Instructions sheet
    # ─────────────────────────────────────────────────────
    print("\n[3] Creating Instructions sheet...")

    try:
        wb.Sheets("Instructions").Delete()
    except:
        pass

    ws = wb.Sheets.Add(Before=wb.Sheets(1))
    ws.Name = "Instructions"
    ws.Tab.Color = xlc("#95A5A6")

    def sc(row, col, val, bold=False, size=11, bg=None, fg=None, italic=False,
           align=-4131, merge_end_col=None):
        cell = ws.Cells(row, col)
        cell.Value = val
        cell.Font.Name = "Calibri"
        cell.Font.Size = size
        cell.Font.Bold = bold
        cell.Font.Italic = italic
        if bg: cell.Interior.Color = xlc(bg)
        if fg: cell.Font.Color = xlc(fg)
        cell.HorizontalAlignment = align
        cell.VerticalAlignment = -4108
        cell.WrapText = True
        if merge_end_col:
            ws.Range(ws.Cells(row, col), ws.Cells(row, merge_end_col)).Merge()

    def section(row, text):
        sc(row, 1, text, bold=True, size=12, bg="#1A3A5C", fg="#FFFFFF",
           align=-4108, merge_end_col=8)
        ws.Rows(row).RowHeight = 24

    def body(row, col, text, bold=False, bg="#FFFFFF"):
        sc(row, col, text, bold=bold, bg=bg, fg="#222222",
           align=-4131, merge_end_col=8)
        ws.Rows(row).RowHeight = 18

    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 60
    for c in range(3, 9): ws.Columns(c).ColumnWidth = 8

    r = 1
    sc(r, 1, "FOREX TRADING WORKBOOK — USER GUIDE",
       bold=True, size=18, bg="#0D2137", fg="#FFFFFF",
       align=-4108, merge_end_col=8)
    ws.Rows(r).RowHeight = 44
    r += 1
    sc(r, 1, "Version 1.0  |  Ready for distribution",
       size=9, bg="#1A3A5C", fg="#AAAAAA",
       italic=True, align=-4108, merge_end_col=8)
    ws.Rows(r).RowHeight = 16
    r += 2

    section(r, "SECTION 1 — COLOR GUIDE"); r += 1
    sc(r, 1, "Blue cells  = YOU enter data here — click and type freely",
       bg="#D6E4F0", fg="#1A3A5C", bold=True, merge_end_col=8)
    ws.Rows(r).RowHeight = 18; r += 1
    sc(r, 1, "Grey cells  = Auto-calculated — do NOT edit",
       bg="#F5F5F5", fg="#666666", merge_end_col=8)
    ws.Rows(r).RowHeight = 18; r += 2

    section(r, "SECTION 2 — THE 5 TABS (what each one does)"); r += 1
    tabs = [
        ("Instructions", "This page. Read first."),
        ("Summary", "Monthly performance dashboard. Fill 6 cells at start of month."),
        ("Daily Log", "Daily trading journal. Fill 4 columns per day."),
        ("Raw daily data", "Paste broker export here daily. Columns A–H only."),
        ("Advanced Setup Planner", "Fill every morning before your session."),
    ]
    for tab, desc in tabs:
        sc(r, 1, tab, bold=True, bg="#F9FAFB", fg="#1A3A5C")
        sc(r, 2, desc, bg="#F9FAFB", fg="#333333", merge_end_col=8)
        ws.Rows(r).RowHeight = 18; r += 1
    r += 1

    section(r, "SECTION 3 — SUMMARY TAB"); r += 1
    body(r, 1, "Fill these 6 blue cells at the start of each month:", bold=True); r += 1
    fields = [
        "Monthly Profit Target ($)     → How much you aim to make",
        "Planned Active Trading Days   → Days you plan to trade",
        "Weeks in Cycle                → Usually 4",
        "Risk % per trade              → e.g. 0.07 = 7% of NAV",
        "Max Daily Loss (R multiples)  → e.g. 2 = stop after losing 2R",
        "Current NAV ($)               → Account balance on day 1",
    ]
    for f in fields:
        sc(r, 1, f, bg="#D6E4F0", fg="#1A3A5C", merge_end_col=8)
        ws.Rows(r).RowHeight = 18; r += 1
    r += 1

    section(r, "SECTION 4 — DAILY LOG TAB"); r += 1
    body(r, 1, "Each trading day, fill ONLY these 4 columns:", bold=True); r += 1
    dl = [
        "Column A → Today's date",
        "Column B → Week number in cycle (1, 2, 3 or 4)",
        "Column D → Y = traded today, N = did not trade",
        "Column M → Notes (psychology, lessons, market conditions)",
    ]
    for d in dl:
        sc(r, 1, d, bg="#D6E4F0", fg="#1A3A5C", merge_end_col=8)
        ws.Rows(r).RowHeight = 18; r += 1
    body(r, 1,
         "Do NOT touch columns E, F, G, H, I, J, K, L, N–Y — they auto-calculate.",
         bold=True, bg="#FFF3CD"); r += 1
    r += 1

    section(r, "SECTION 5 — RAW DAILY DATA TAB"); r += 1
    steps5 = [
        "Step 1: In MT4/MT5 → Account History → right-click → Save as Report",
        "Step 2: Copy columns: Symbol | Name | Side | Quantity | Price | Profit | Deal ID | Time",
        "Step 3: Delete rows 2+ in 'Raw daily data', paste starting at A2",
        "NOTE:   Column H (Time) must be: \"Wed, 23 Apr 2026 09:00:00 GMT\" format",
        "NOTE:   Columns I and J are formulas — never overwrite them",
    ]
    for s in steps5:
        sc(r, 1, s, bg="#FFFFFF" if not s.startswith("NOTE") else "#FFF3CD",
           fg="#222222", merge_end_col=8)
        ws.Rows(r).RowHeight = 18; r += 1
    r += 1

    section(r, "SECTION 6 — ADVANCED SETUP PLANNER TAB"); r += 1
    body(r, 1, "Fill this every morning BEFORE your trading session.", bold=True); r += 1
    steps6 = [
        ("Step 1 — EMAs", "EMA9/21/50 on 15m, SMA200 on 15m, EMA50 on 1H/4H/1D/1W, EMA100/200 on 1D, EMA20 on 1M, Current Price"),
        ("Step 2 — RSI & ATR", "RSI(14) from 15m, 1H, 4H. ATR(14) from 15m. ADR(14)."),
        ("Step 3 — Price Levels", "PWH/PWL, PDH/PDL, EQH/EQL, LTF Leg, Session High/Low, HTF Swing High/Low."),
        ("Step 4 — FVG Zones", "Daily FVG high/low, 1H FVG high/low. Leave blank if no clear FVG."),
        ("Step 5 — Options Gamma", "Up to 4 strike prices with open interest (bn). Leave blank if unavailable."),
        ("Step 6 — Outputs", "SIDE, TRADE MODE, 4 LADDERS, NARRATIVE are auto-calculated — do not edit."),
    ]
    for label, desc in steps6:
        sc(r, 1, label, bold=True, bg="#F9FAFB", fg="#1A3A5C")
        sc(r, 2, desc, bg="#FFFFFF", fg="#333333", merge_end_col=8)
        ws.Rows(r).RowHeight = 28; r += 1
    r += 1

    section(r, "SECTION 7 — NEW MONTH RESET"); r += 1
    resets = [
        "1. Save current file as Trading_[Month][Year].xlsx (e.g. Trading_Apr2026.xlsx)",
        "2. Open this master template",
        "3. Update Summary → Current NAV = ending balance of last month",
        "4. Clear Daily Log columns A, B, D, M (keep formulas in other columns)",
        "5. Delete rows 2+ in Raw daily data",
        "6. Clear all blue input cells in Advanced Setup Planner",
        "7. Save — ready for new month",
    ]
    for s in resets:
        body(r, 1, s); r += 1
    r += 1

    section(r, "SECTION 8 — FAQs"); r += 1
    faqs = [
        ('Q: Some cells show "" (empty) — is that a bug?',
         'A: No. Cells are intentionally blank until you enter data.'),
        ('Q: CT Ladder shows "CT inactive"?',
         'A: CT trades only apply when SIDE = SHORT. When LONG, CT is disabled.'),
        ('Q: Macro Bias shows "Neutral" — should I trade?',
         'A: Neutral = no clear macro edge. If Trade Mode = Avoid, do not trade.'),
        ('Q: Can I add more rows to Daily Log?',
         'A: Yes. Insert a row, copy formulas from the row above.'),
    ]
    for q, a in faqs:
        body(r, 1, q, bold=True, bg="#FFF9E6"); r += 1
        body(r, 1, a); r += 1
    r += 1

    # Navigation hyperlinks
    section(r, "NAVIGATION"); r += 1
    nav = [
        ("→ Go to Summary", "Summary!A1"),
        ("→ Go to Daily Log", "'Daily Log'!A1"),
        ("→ Go to Raw daily data", "'Raw daily data'!A1"),
        ("→ Go to Advanced Setup Planner", "'Advanced Setup Planner'!A1"),
    ]
    for i, (label, target) in enumerate(nav):
        cell = ws.Cells(r, 1 + i * 2)
        ws.Hyperlinks.Add(Anchor=cell, Address="", SubAddress=target,
                          TextToDisplay=label)
        cell.Font.Color = xlc("#1A7A4A")
        cell.Font.Underline = 2
        cell.Font.Bold = True
        cell.Font.Name = "Calibri"
        cell.Font.Size = 11

    print("    Instructions sheet created.")

    # ─────────────────────────────────────────────────────
    # STEP 5: Backend → VeryHidden
    # ─────────────────────────────────────────────────────
    print("\n[4] Setting Backend to VeryHidden...")
    try:
        wb.Sheets("Backend").Visible = 2  # xlSheetVeryHidden
        print("    Backend: VeryHidden ✓")
    except Exception as e:
        print(f"    WARNING: {e}")

    # ─────────────────────────────────────────────────────
    # SAVE
    # ─────────────────────────────────────────────────────
    print("\nSaving...")
    excel.DisplayAlerts = True
    wb.SaveAs(dst)
    wb.Close()
    excel.Quit()
    print(f"\n{'=' * 60}")
    print(f"DONE! Saved: Trading_Workbook_FINAL.xlsx")
    print(f"{'=' * 60}")

if __name__ == '__main__':
    main()
