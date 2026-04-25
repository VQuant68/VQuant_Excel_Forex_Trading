"""Auto check steps 7 through 13."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

print("\n=== CHECK 7 & 8: Cell Protection ===")
sh_planner = wb.Sheets("Advanced Setup Planner")
# Formula cell (B31) should be Locked = True
formula_locked = sh_planner.Range("B31").Locked
# Input cell (B5) should be Locked = False
input_locked = sh_planner.Range("B5").Locked
# Sheet should be protected
is_protected = sh_planner.ProtectContents

if is_protected and formula_locked and not input_locked:
    print("✅ PASS: Formula cells are locked, input cells are unlocked, sheet is protected.")
else:
    print(f"❌ FAIL: Protected={is_protected}, B31 Locked={formula_locked}, B5 Locked={input_locked}")

print("\n=== CHECK 9: Backend VeryHidden ===")
try:
    sh_backend = wb.Sheets("Backend")
    vis = sh_backend.Visible
    if vis == 2:  # 2 = xlSheetVeryHidden
        print("✅ PASS: Backend is VeryHidden (not in Unhide list).")
    else:
        print(f"❌ FAIL: Backend visibility is {vis}")
except Exception as e:
    print(f"❌ FAIL: {e}")

print("\n=== CHECK 10: Navigation Buttons ===")
sh_summary = wb.Sheets("Summary")
if sh_summary.Shapes.Count > 0:
    print("✅ PASS: Navigation shapes found on Summary.")
else:
    print("❌ FAIL: No shapes found.")

print("\n=== CHECK 11: No Error Values ===")
error_count = 0
for sn in ["Summary", "Daily Log", "Raw daily data", "Advanced Setup Planner"]:
    sh = wb.Sheets(sn)
    used = sh.UsedRange
    for r in range(1, min(used.Rows.Count + 1, 50)):
        for c in range(1, min(used.Columns.Count + 1, 30)):
            try:
                txt = str(sh.Cells(r, c).Text)
                if txt in ["#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?"]:
                    error_count += 1
            except: pass
if error_count == 0:
    print("✅ PASS: No #REF, #DIV/0, #VALUE errors found.")
else:
    print(f"❌ FAIL: Found {error_count} error cells.")

print("\n=== CHECK 12: Format $ and % ===")
sh_sum = wb.Sheets("Summary")
fmt_b10 = sh_sum.Range("B10").NumberFormat
sh_dl = wb.Sheets("Daily Log")
fmt_s2 = sh_dl.Range("S2").NumberFormat

print(f"   Summary B10 format: '{fmt_b10}'")
print(f"   Daily Log S2 format: '{fmt_s2}'")
if "$" in fmt_b10 and "%" in fmt_s2:
    print("✅ PASS: Formats look correct.")
else:
    print("⚠️ FORMAT MAY BE WRONG. Check visually.")

print("\n=== CHECK 13: Conditional Formatting (P&L) ===")
fc_count = sh_dl.Range("H2").FormatConditions.Count
if fc_count > 0:
    print(f"✅ PASS: Found {fc_count} conditional formatting rules on P&L column.")
else:
    print("❌ FAIL: No conditional formatting rules on H2.")

wb.Close(False); excel.Quit()
