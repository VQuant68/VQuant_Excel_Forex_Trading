"""CHECK 3-6: Inject EMAs → verify Side, Ladder, Narrative, Macro Bias."""
import win32com.client, os, time

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass

# Bullish EMAs (Price above all EMAs = LONG)
data = {
    5: 1.17, 6: 1.1695, 7: 1.169, 8: 1.1685, 9: 1.165,
    10: 1.168, 11: 1.167, 12: 1.166, 13: 1.164, 14: 1.15,
    15: 1.14, 16: 1.13, 17: 58, 18: 55, 19: 60, 20: 0.0008, 21: 80
}
for r, v in data.items():
    sh.Cells(r, 2).Value = v

# Leg Low/High for ladders
sh.Cells(13, 5).Value = 1.1600
sh.Cells(14, 5).Value = 1.1700

excel.CalculateFullRebuild()
time.sleep(1)

# ── CHECK 3 ──
side = sh.Range("B31").Value
print(f"✅ CHECK 3: Side = '{side}'")
print(f"   {'PASS ✅' if side and 'LONG' in str(side).upper() else 'FAIL ❌'}")

# ── CHECK 4 ──
setup = sh.Range("B32").Value
print(f"\n✅ CHECK 4: Setup = '{str(setup)[:80]}'")
print(f"   {'PASS ✅' if setup else 'FAIL ❌'}")

# ── CHECK 5 ──
narrative = sh.Range("B38").Value
print(f"\n✅ CHECK 5: Narrative = '{str(narrative)[:80]}...'")
print(f"   {'PASS ✅' if narrative and len(str(narrative)) > 10 else 'FAIL ❌'}")

# ── CHECK 6 ──
bias = sh.Range("B34").Value
print(f"\n✅ CHECK 6: Macro Bias = '{bias}'")
print(f"   {'PASS ✅' if bias else 'FAIL ❌'}")

# ── CHECK 33 ──
trade_mode = sh.Range("B33").Value
print(f"\n   Trade Mode = '{trade_mode}'")

# Re-protect (keep data for user to see)
sh.Range("B5:B21").Locked = False
sh.Range("E5:E18").Locked = False
sh.Range("C24:C27").Locked = False
sh.Range("D24:D27").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("\n=== MỞ FILE XEM PLANNER ===")
