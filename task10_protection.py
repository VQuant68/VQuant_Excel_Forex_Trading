"""
VIỆC 10: Sheet Protection - Lock formulas, Unlock inputs, Protect all sheets.
Password: cuongdeptrai
"""
import win32com.client, os

PWD = "cuongdeptrai"

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

def protect_sheet(ws, unlock_ranges=None):
    """Lock all → Unlock specified ranges → Protect sheet."""
    name = ws.Name
    
    # Unprotect first if already protected
    try: ws.Unprotect(PWD)
    except: pass
    try: ws.Unprotect()
    except: pass
    
    # 1. Lock ALL cells
    ws.Cells.Locked = True
    
    # 2. Unlock input ranges
    if unlock_ranges:
        for rng_addr in unlock_ranges:
            try:
                rng = ws.Range(rng_addr)
                rng.Locked = False
                print(f"  Unlocked: {rng_addr}")
            except Exception as e:
                print(f"  WARNING: Could not unlock {rng_addr}: {e}")
    
    # 3. Protect sheet
    # AllowSelectLockedCells, AllowSelectUnlockedCells, AllowAutoFilter
    ws.Protect(
        Password=PWD,
        DrawingObjects=True,
        Contents=True,
        Scenarios=True,
        AllowFormattingCells=False,
        AllowFormattingColumns=False,
        AllowFormattingRows=False,
        AllowInsertingColumns=False,
        AllowInsertingRows=False,
        AllowInsertingHyperlinks=False,
        AllowDeletingColumns=False,
        AllowDeletingRows=False,
        AllowSorting=False,
        AllowFiltering=True,       # Allow AutoFilter
        AllowUsingPivotTables=False,
    )
    # These are set after Protect
    ws.EnableSelection = 0  # 0=xlNoRestrictions (select both locked/unlocked)
    
    print(f"✅ {name}: Protected (password set)")

# ══════════════════════════════════════════════════════
# INSTRUCTIONS — Lock everything
# ══════════════════════════════════════════════════════
print("\n── Instructions ──")
protect_sheet(wb.Sheets("Instructions"))

# ══════════════════════════════════════════════════════
# SUMMARY — Unlock B2, B3, B5, B8, B9, B10
# ══════════════════════════════════════════════════════
print("\n── Summary ──")
protect_sheet(wb.Sheets("Summary"), [
    "B2", "B3", "B5", "B8", "B9", "B10"
])

# ══════════════════════════════════════════════════════
# DAILY LOG — Unlock A2:A100, B2:B100, D2:D100, M2:M100
# ══════════════════════════════════════════════════════
print("\n── Daily Log ──")
protect_sheet(wb.Sheets("Daily Log"), [
    "A2:A100", "B2:B100", "D2:D100", "M2:M100"
])

# ══════════════════════════════════════════════════════
# RAW DAILY DATA — Unlock A2:H500
# ══════════════════════════════════════════════════════
print("\n── Raw daily data ──")
protect_sheet(wb.Sheets("Raw daily data"), [
    "A2:H500"
])

# ══════════════════════════════════════════════════════
# ADVANCED SETUP PLANNER — Unlock all blue input cells
# ══════════════════════════════════════════════════════
print("\n── Advanced Setup Planner ──")
planner_inputs = [
    # EMAs + indicators (B5:B21)
    "B5:B21",
    # Price levels (E5:E18) 
    "E5:E18",
    # BOS/CHOCH manual inputs (C24:D27)
    "C24:C27", "D24:D27",
]

# Also find engine param inputs (Magnet, Weights, etc.)
sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass
try: sh.Unprotect()
except: pass

# Scan for input cells with blue background in cols H-K
for r in range(1, 50):
    for c in range(8, 12):  # H=8, I=9, J=10, K=11
        try:
            cell = sh.Cells(r, c)
            bg = cell.Interior.Color
            # Check if it's an input (blue #D6E4F0 = 14738646 in BGR)
            if bg == 14738646 or bg == 15787734:  # possible blue variants
                addr = cell.Address.replace("$", "")
                planner_inputs.append(addr)
        except:
            pass

# FVG zones in col N area
for r in range(1, 50):
    for c in range(13, 20):  # M-T
        try:
            cell = sh.Cells(r, c)
            bg = cell.Interior.Color
            if bg == 14738646 or bg == 15787734:
                addr = cell.Address.replace("$", "")
                planner_inputs.append(addr)
        except:
            pass

protect_sheet(wb.Sheets("Advanced Setup Planner"), planner_inputs)

# ══════════════════════════════════════════════════════
# BACKEND — Don't protect (will be hidden in Việc 11)
# ══════════════════════════════════════════════════════
print("\n── Backend ── (skipped, will be hidden)")

# ══════════════════════════════════════════════════════
# WORKBOOK PROTECTION — Prevent sheet deletion/renaming
# ══════════════════════════════════════════════════════
print("\n── Workbook Structure Protection ──")
wb.Protect(Password=PWD, Structure=True, Windows=False)
print("✅ Workbook structure protected")

wb.Save()
wb.Close()
excel.Quit()
print("\n=== VIỆC 10 HOÀN THÀNH ===")
print("Test: open file, click formula cell → should show 'protected' message")
print("Test: click blue input cell → should allow editing")
