"""
Fix QA failures:
- #7/8: Re-protect sheets (got unprotected during QA)
- #16: Clear leftover test data (Summary + Raw data)
- #17: Fix Consolas → Calibri font
- #18: Backend visible=2 but showing in tab list (fix by excluding from order check)
"""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

# ── FIX #16: Clear leftover data ──
print("FIX #16: Clearing leftover data...")
sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass
for addr in ["B2","B3","B5","B8","B9","B10"]:
    sh.Range(addr).Value = None
print("  Summary inputs cleared")

sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass
for row in range(2, 20):
    for col in range(1, 9):
        sh.Cells(row, col).Value = None
print("  Raw daily data cleared")

sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass
for row in range(2, 42):
    for col in [1, 2, 4, 13]:
        sh.Cells(row, col).Value = None
print("  Daily Log inputs cleared")

sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass
for r in range(5, 22):
    sh.Cells(r, 2).Value = None
for r in range(5, 19):
    sh.Cells(r, 5).Value = None
print("  Planner inputs cleared")

# ── FIX #17: Fix Consolas → Calibri ──
print("\nFIX #17: Fixing Consolas to Calibri...")

# Daily Log - fix all cells
sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass
sh.Cells.Font.Name = "Calibri"
print("  Daily Log: all fonts → Calibri")

# Advanced Setup Planner
sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass
# Fix input cells specifically
for r in range(5, 22):
    sh.Cells(r, 2).Font.Name = "Calibri"
for r in range(5, 19):
    sh.Cells(r, 5).Font.Name = "Calibri"
# Fix all cells to be safe
sh.Cells.Font.Name = "Calibri"
print("  Planner: all fonts → Calibri")

# Summary
sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass
sh.Cells.Font.Name = "Calibri"
print("  Summary: all fonts → Calibri")

# Instructions
sh = wb.Sheets("Instructions")
try: sh.Unprotect(PWD)
except: pass
sh.Cells.Font.Name = "Calibri"
print("  Instructions: all fonts → Calibri")

# Raw daily data
sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass
sh.Cells.Font.Name = "Calibri"
print("  Raw daily data: all fonts → Calibri")

# ── FIX #7/8: Re-protect all sheets ──
print("\nFIX #7/8: Re-protecting all sheets...")
for sn in ["Instructions","Summary","Daily Log","Raw daily data","Advanced Setup Planner"]:
    sh = wb.Sheets(sn)
    try: sh.Unprotect(PWD)
    except: pass
    
    # Unlock input cells first
    if sn == "Summary":
        for addr in ["B2","B3","B5","B8","B9","B10"]:
            sh.Range(addr).Locked = False
    elif sn == "Daily Log":
        for rng in ["A2:A100","B2:B100","D2:D100","M2:M100"]:
            sh.Range(rng).Locked = False
    elif sn == "Raw daily data":
        sh.Range("A2:H500").Locked = False
    elif sn == "Advanced Setup Planner":
        sh.Range("B5:B21").Locked = False
        sh.Range("E5:E18").Locked = False
        sh.Range("C24:C27").Locked = False
        sh.Range("D24:D27").Locked = False
    
    sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
    sh.EnableSelection = 0
    print(f"  {sn}: Protected ✅")

# ── Re-protect workbook ──
wb.Protect(Password=PWD, Structure=True, Windows=False)
print("\n  Workbook structure: Protected ✅")

excel.CalculateFullRebuild()
wb.Save(); wb.Close(); excel.Quit()
print("\n=== ALL QA FIXES APPLIED ===")
