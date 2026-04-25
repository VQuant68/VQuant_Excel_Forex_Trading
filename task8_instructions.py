"""
VIỆC 8 v2: Professional Instructions sheet — clean, spacious layout
All content merged B:K for wide readable text.
"""
import win32com.client, os

def xlc(h):
    h=h.lstrip('#'); return int(h[0:2],16)+int(h[2:4],16)*256+int(h[4:6],16)*65536

C_TITLE_BG = xlc("#0D2137")
C_TITLE_FG = xlc("#FFFFFF")
C_HDR_BG   = xlc("#1A3A5C")
C_HDR_FG   = xlc("#FFFFFF")
C_INPUT_BG = xlc("#D6E4F0")
C_INPUT_FG = xlc("#1A3A5C")
C_FORM_BG  = xlc("#F5F5F5")
C_FORM_FG  = xlc("#666666")
C_WHITE    = xlc("#FFFFFF")
C_LIGHT    = xlc("#F9FAFB")
C_BORDER   = xlc("#E0E0E0")
C_TEXT     = xlc("#333333")
C_WARN     = xlc("#B71C1C")

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

try: wb.Sheets("Instructions").Delete()
except: pass

ws = wb.Sheets.Add(Before=wb.Sheets(1))
ws.Name = "Instructions"
ws.Tab.Color = xlc("#95A5A6")

# ─── Setup columns ───────────────────────────────────
ws.Columns("A").ColumnWidth = 3     # left margin
ws.Columns("B").ColumnWidth = 14    # label/number
ws.Columns("C").ColumnWidth = 14
ws.Columns("D").ColumnWidth = 14
ws.Columns("E").ColumnWidth = 14
ws.Columns("F").ColumnWidth = 14
ws.Columns("G").ColumnWidth = 14
ws.Columns("H").ColumnWidth = 14
ws.Columns("I").ColumnWidth = 14
ws.Columns("J").ColumnWidth = 14
ws.Columns("K").ColumnWidth = 14
ws.Columns("L").ColumnWidth = 3     # right margin

# White bg for entire area
ws.Range("A1:L200").Interior.Color = C_WHITE
ws.Range("A1:L200").Font.Name = "Calibri"
ws.Range("A1:L200").Font.Size = 11

r = 1

# ─── Helpers ──────────────────────────────────────────
def merge_write(row, text, bg=None, fg=C_TEXT, bold=False, size=11, 
                height=22, halign=-4131, italic=False, col_start="B", col_end="K"):
    rng = ws.Range(f"{col_start}{row}:{col_end}{row}")
    try: rng.UnMerge()
    except: pass
    try: rng.Merge()
    except: pass
    ws.Cells(row, ord(col_start[0])-64).Value = text
    if bg: rng.Interior.Color = bg
    rng.Font.Color = fg
    rng.Font.Bold = bold
    rng.Font.Size = size
    rng.Font.Italic = italic
    rng.Font.Name = "Calibri"
    rng.HorizontalAlignment = halign  # -4131=left, -4108=center
    rng.VerticalAlignment = -4108     # center
    rng.WrapText = True
    ws.Rows(row).RowHeight = height

def title(row, text):
    merge_write(row, text, bg=C_TITLE_BG, fg=C_TITLE_FG, bold=True, size=16, height=44, halign=-4108)

def section(row, text):
    merge_write(row, text, bg=C_HDR_BG, fg=C_HDR_FG, bold=True, size=12, height=30)

def body(row, text, bold=False, italic=False, fg=C_TEXT, bg=None, size=11, height=20):
    merge_write(row, text, fg=fg, bold=bold, italic=italic, bg=bg, size=size, height=height)

def spacer(row, h=6):
    ws.Rows(row).RowHeight = h

def bullet(row, label, desc, label_bg=None, label_fg=C_TEXT):
    """Two-cell row: B:D = label, E:K = description"""
    rng_l = ws.Range(f"B{row}:D{row}")
    try: rng_l.UnMerge()
    except: pass
    try: rng_l.Merge()
    except: pass
    ws.Cells(row, 2).Value = "  " + label
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 11
    ws.Cells(row, 2).Font.Name = "Calibri"
    ws.Cells(row, 2).Font.Color = label_fg
    ws.Cells(row, 2).HorizontalAlignment = -4131
    ws.Cells(row, 2).VerticalAlignment = -4108
    if label_bg: rng_l.Interior.Color = label_bg
    
    rng_r = ws.Range(f"E{row}:K{row}")
    try: rng_r.UnMerge()
    except: pass
    try: rng_r.Merge()
    except: pass
    ws.Cells(row, 5).Value = desc
    ws.Cells(row, 5).Font.Size = 11
    ws.Cells(row, 5).Font.Name = "Calibri"
    ws.Cells(row, 5).Font.Color = C_TEXT
    ws.Cells(row, 5).HorizontalAlignment = -4131
    ws.Cells(row, 5).VerticalAlignment = -4108
    ws.Rows(row).RowHeight = 22

def thin_border_bottom(row):
    ws.Range(f"B{row}:K{row}").Borders(9).LineStyle = 1  # xlBottom
    ws.Range(f"B{row}:K{row}").Borders(9).Color = C_BORDER
    ws.Range(f"B{row}:K{row}").Borders(9).Weight = 1

# ══════════════════════════════════════════════════════
# BUILD CONTENT
# ══════════════════════════════════════════════════════

# ── TITLE ──
title(r, "FOREX TRADING WORKBOOK — USER GUIDE"); r += 1
merge_write(r, "Version 1.0  |  Ready for distribution", fg=C_FORM_FG, italic=True, size=10, height=24, halign=-4108)
r += 1; spacer(r, 12); r += 1

# ── SECTION 1: COLOR GUIDE ──
section(r, "  SECTION 1 — COLOR GUIDE"); r += 1
spacer(r, 8); r += 1
bullet(r, "■  INPUT CELL", "YOU enter data here — these are your inputs", label_bg=C_INPUT_BG, label_fg=C_INPUT_FG)
r += 1
bullet(r, "■  FORMULA CELL", "Auto-calculated by the workbook — do NOT edit", label_bg=C_FORM_BG, label_fg=C_FORM_FG)
r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 2: HOW IT WORKS ──
section(r, "  SECTION 2 — HOW THIS WORKBOOK WORKS"); r += 1
spacer(r, 8); r += 1
body(r, "  This workbook has 5 tabs:", bold=True); r += 1
spacer(r, 4); r += 1

tabs = [
    ("1.  Instructions", "You are here. Read this first."),
    ("2.  Summary", "Monthly performance dashboard. Fill once per month."),
    ("3.  Daily Log", "Your daily trading journal. Fill each trading day."),
    ("4.  Raw daily data", "Paste your broker trade history here daily."),
    ("5.  Setup Planner", "Fill before each trading session for auto trade plan."),
]
for name, desc in tabs:
    bullet(r, name, desc, label_fg=C_HDR_BG); r += 1

spacer(r, 8); r += 1
body(r, "  Data Flow:", bold=True, fg=C_HDR_BG); r += 1
body(r, "      Raw daily data  ➜  Daily Log  ➜  Summary          (automatic aggregation)", fg=xlc("#2E86AB"), bold=True, size=10); r += 1
body(r, "      Your inputs  ➜  Setup Planner  ➜  Trade Plan     (automatic calculation)", fg=xlc("#E67E22"), bold=True, size=10); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 3: SUMMARY ──
section(r, "  SECTION 3 — SUMMARY TAB"); r += 1
spacer(r, 8); r += 1
body(r, "  Fill these 6 cells at the start of each month:", bold=True); r += 1
spacer(r, 4); r += 1

s_items = [
    ("Monthly Target ($)", "How much you aim to make this month"),
    ("Active Trading Days", "How many days you plan to trade"),
    ("Weeks in Cycle", "Usually 4"),
    ("Risk % per trade", "e.g. 0.07 = 7% of NAV"),
    ("Max Daily Loss (R)", "e.g. 2 = stop after -2R intraday"),
    ("Current NAV ($)", "Your account balance on day 1"),
]
for name, desc in s_items:
    bullet(r, name, desc, label_bg=C_INPUT_BG, label_fg=C_INPUT_FG); r += 1

spacer(r, 6); r += 1
body(r, "  Everything else auto-calculates:", bold=True, fg=C_FORM_FG); r += 1
auto = [
    ("Projected P&L", "Forecast if you maintain current pace"),
    ("Required R/Day", "How many R you need each remaining day"),
    ("Weekly Table", "Auto-summary for each of the 4 weeks"),
]
for name, desc in auto:
    bullet(r, name, desc, label_bg=C_FORM_BG, label_fg=C_FORM_FG); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 4: DAILY LOG ──
section(r, "  SECTION 4 — DAILY LOG TAB"); r += 1
spacer(r, 8); r += 1
body(r, "  Each trading day, fill only these 4 columns:", bold=True); r += 1
spacer(r, 4); r += 1

dl = [
    ("Column A", "Today's date"),
    ("Column B", "Week of the cycle (1, 2, 3, or 4)"),
    ("Column D", "Y if you traded, N if not"),
    ("Column M", "Notes — psychology, conditions, lessons"),
]
for name, desc in dl:
    bullet(r, name, desc, label_bg=C_INPUT_BG, label_fg=C_INPUT_FG); r += 1

spacer(r, 6); r += 1
body(r, "  ⚠  Do NOT touch columns E–L, N–Y — they auto-calculate.", bold=True, fg=C_WARN); r += 1
spacer(r, 4); r += 1
body(r, "  Column reference:", bold=True, size=10, fg=C_FORM_FG); r += 1
body(r, "  E=NAV | F=Risk$ | G=MaxLoss | H=P&L | I=CumPnL | J=WeekPnL | K=Target | L=Variance", size=9, fg=C_FORM_FG); r += 1
body(r, "  N=#Trades | O=#Wins | P=#Loss | Q=Win$ | R=Loss$ | S=WinRate | T=R:R | U=DailyR | V=Loss%", size=9, fg=C_FORM_FG); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 5: RAW DATA ──
section(r, "  SECTION 5 — RAW DAILY DATA TAB"); r += 1
spacer(r, 8); r += 1
body(r, "  How to import from your broker (MT4/MT5):", bold=True); r += 1
spacer(r, 4); r += 1

steps5 = [
    "Step 1:   In MT4/MT5, go to Account History",
    "Step 2:   Right-click → Save as Report → select date range",
    "Step 3:   Open exported file, copy: Symbol | Name | Side | Qty | Price | Profit | DealID | Time",
    "Step 4:   Delete rows 2+ in 'Raw daily data' tab",
    "Step 5:   Paste data starting from row 2, columns A–H",
]
for s in steps5:
    body(r, "      " + s, size=10); r += 1

spacer(r, 6); r += 1
body(r, "  ⚠  IMPORTANT:", bold=True, fg=C_WARN); r += 1
body(r, '      · Time format must be:  "Wed, 23 Apr 2026 09:00:00 GMT"', fg=C_WARN, size=10); r += 1
body(r, "      · Columns I–J are formulas — never overwrite them", fg=C_WARN, size=10); r += 1
body(r, "      · Paste in chronological order (oldest first)", fg=C_WARN, size=10); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 6: PLANNER ──
section(r, "  SECTION 6 — ADVANCED SETUP PLANNER TAB"); r += 1
spacer(r, 8); r += 1
body(r, "  Fill this every morning before your trading session.", bold=True, italic=True); r += 1
spacer(r, 6); r += 1

psteps = [
    ("Step 1 — EMAs", "EMA 9/21/50 (15m), SMA200 (15m), EMA50 (1H/4H), EMA 50/100/200 (1D), EMA50 (1W), EMA20 (1M), Current Price"),
    ("Step 2 — RSI & Vol", "RSI(14) from 15m, 1H, 4H  |  ATR(14) 15m  |  ADR(14)"),
    ("Step 3 — Levels", "PWH/PWL, PDH/PDL, EQH/EQL, 15m Swing Leg, Session H/L, 4H HTF Swing H/L"),
    ("Step 4 — FVG Zones", "Daily FVG high/low, 1H FVG high/low (leave blank if none)"),
    ("Step 5 — Gamma", "Up to 4 strikes + size in billions (optional)"),
]
for name, desc in psteps:
    bullet(r, name, desc, label_fg=C_HDR_BG); r += 1

spacer(r, 8); r += 1
body(r, "  Step 6 — Read the Outputs (auto-calculated):", bold=True, fg=C_HDR_BG); r += 1
spacer(r, 4); r += 1
outs = [
    ("SIDE", "LONG or SHORT — based on HTF EMA alignment"),
    ("TRADE MODE", "CoreLong / A+Long / CTshort / Avoid"),
    ("LADDERS", "Primary / Value / A+ / CT — Fibonacci entry rungs"),
    ("NARRATIVE", "Plain English macro + structure summary"),
]
for name, desc in outs:
    bullet(r, name, desc, label_bg=C_FORM_BG, label_fg=C_FORM_FG); r += 1

spacer(r, 8); r += 1
body(r, "  Trade Mode explained:", bold=True); r += 1
spacer(r, 4); r += 1
modes = [
    ("CoreLong", "HTF trend aligned — standard continuation trade", "#C6EFCE", "#006100"),
    ("A+Long", "Premium deep zone — can size up slightly", "#C6EFCE", "#006100"),
    ("CTshort", "Counter-trend fade — reduce size", "#FFC7CE", "#9C0006"),
    ("Avoid", "Conditions not met — do NOT trade today", "#FF9999", "#7F0000"),
]
for mode, desc, bg, fg in modes:
    bullet(r, mode, desc, label_bg=xlc(bg), label_fg=xlc(fg)); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 7: MACRO ──
section(r, "  SECTION 7 — MACRO DATA (advanced users)"); r += 1
spacer(r, 8); r += 1
body(r, "  The macro scoring engine is pre-built. Update weekly on Mondays:", size=10); r += 1
body(r, "      · Add economic events (CPI, NFP, GDP, rate decisions)", size=10); r += 1
body(r, "      · Add central bank remarks: Hawkish / Neutral / Dovish", size=10); r += 1
body(r, "      → Contact your account manager for backend access.", italic=True, fg=C_FORM_FG, size=10); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 8: RESET ──
section(r, "  SECTION 8 — NEW MONTH RESET"); r += 1
spacer(r, 8); r += 1
body(r, "  At the start of each new month:", bold=True); r += 1
spacer(r, 4); r += 1
resets = [
    "1.  Save current file as Trading_[Month][Year].xlsx",
    "2.  Open master template (this file)",
    "3.  Update Summary → Current NAV = last month's ending balance",
    "4.  Update Summary → Monthly Target if changed",
    "5.  Delete Daily Log columns A, B, D, M content (keep formulas)",
    "6.  Delete rows 2+ in Raw daily data",
    "7.  Clear all blue input cells in Advanced Setup Planner",
    "8.  Save — ready for the new month ✓",
]
for s in resets:
    body(r, "      " + s, size=10); r += 1
thin_border_bottom(r-1); spacer(r, 12); r += 1

# ── SECTION 9: FAQ ──
section(r, "  SECTION 9 — FREQUENTLY ASKED QUESTIONS"); r += 1
spacer(r, 8); r += 1

faqs = [
    ('Q:  Cells show "" (empty) instead of numbers?',
     'A:  Normal. Cells stay blank until you enter data — prevents errors.'),
    ('Q:  Can I add more rows to Daily Log?',
     'A:  Yes. Insert row, then copy formulas from the row above.'),
    ('Q:  CT Ladder shows "CT inactive"?',
     'A:  CT trades only apply when SIDE = SHORT. Disabled when LONG.'),
    ('Q:  Macro Bias = "Neutral" — should I trade?',
     'A:  Neutral = balanced. Check technicals. If Trade Mode = Avoid, skip.'),
]
for q, a in faqs:
    body(r, "  " + q, bold=True, fg=C_HDR_BG, size=10); r += 1
    body(r, "  " + a, italic=True, fg=C_FORM_FG, size=10); r += 1
    spacer(r, 6); r += 1

thin_border_bottom(r-1); spacer(r, 16); r += 1

# ── NAVIGATION BUTTONS ──
section(r, "  QUICK NAVIGATION — CLICK TO JUMP"); r += 1
spacer(r, 10); r += 1

nav = [
    ("B", "D", "→  Summary", "Summary", "#1A3A5C"),
    ("E", "F", "→  Daily Log", "Daily Log", "#2E86AB"),
    ("G", "H", "→  Raw Data", "Raw daily data", "#7B68EE"),
    ("I", "K", "→  Setup Planner", "Advanced Setup Planner", "#E67E22"),
]
for cs, ce, label, sheet, color in nav:
    rng = ws.Range(f"{cs}{r}:{ce}{r}")
    try: rng.UnMerge()
    except: pass
    try: rng.Merge()
    except: pass
    col_idx = ord(cs) - 64
    ws.Cells(r, col_idx).Value = label
    rng.Interior.Color = xlc(color)
    rng.Font.Color = C_WHITE
    rng.Font.Bold = True
    rng.Font.Size = 11
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    ws.Hyperlinks.Add(
        Anchor=rng, Address="", SubAddress=f"'{sheet}'!A1", TextToDisplay=label
    )
ws.Rows(r).RowHeight = 32
r += 1

# ─── Final polish ────────────────────────────────────
ws.Activate()
excel.Visible = True
excel.ActiveWindow.DisplayGridlines = False
excel.ActiveWindow.FreezePanes = False
ws.Range("A2").Select()
excel.ActiveWindow.FreezePanes = True
excel.Visible = False

print(f"Instructions: {r} rows, polished layout!")
wb.Save(); wb.Close(); excel.Quit()
print("=== VIỆC 8 v2 HOÀN THÀNH ===")
