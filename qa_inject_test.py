"""Inject CHECK 1-2 test data automatically."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

# ── Summary inputs ──
sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass
sh.Range("B2").Value = 5000
sh.Range("B3").Value = 20
sh.Range("B5").Value = 4
sh.Range("B8").Value = 0.07
sh.Range("B9").Value = 2
sh.Range("B10").Value = 10000
print("Summary inputs filled")

# ── Raw daily data: 5 trades ──
sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass
trades = [
    ["EURUSD","EURUSD","Buy",0.1,1.1350,25.50,"1001","Wed, 23 Apr 2025 09:00:00 GMT"],
    ["EURUSD","EURUSD","Buy",0.1,1.1355,18.00,"1002","Wed, 23 Apr 2025 10:00:00 GMT"],
    ["EURUSD","EURUSD","Sell",0.1,1.1380,-12.00,"1003","Wed, 23 Apr 2025 11:00:00 GMT"],
    ["EURUSD","EURUSD","Buy",0.1,1.1340,35.00,"1004","Wed, 23 Apr 2025 14:00:00 GMT"],
    ["EURUSD","EURUSD","Sell",0.1,1.1390,-8.50,"1005","Wed, 23 Apr 2025 15:00:00 GMT"],
]
for i, trade in enumerate(trades):
    for j, val in enumerate(trade):
        sh.Cells(i+2, j+1).Value = val
print("5 trades pasted into Raw daily data")

# ── Daily Log ──
sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass
sh.Range("A2").Value = "4/23/2025"
sh.Range("B2").Value = 1
sh.Range("D2").Value = "Y"
print("Daily Log date filled")

excel.CalculateFullRebuild()

# ── Read results ──
print("\n=== CHECK 1: Daily Log Results ===")
print(f"  H2 (Net P&L)  = {sh.Range('H2').Value}")
print(f"  N2 (#Trades)   = {sh.Range('N2').Value}")
print(f"  O2 (#Wins)     = {sh.Range('O2').Value}")
print(f"  S2 (Win Rate)  = {sh.Range('S2').Value}")
print(f"  T2 (R:R)       = {sh.Range('T2').Value}")
print(f"  U2 (R-mult)    = {sh.Range('U2').Value}")

print("\n=== CHECK 2: Summary Results ===")
sh_sum = wb.Sheets("Summary")
print(f"  B11 (Total P&L)     = {sh_sum.Range('B11').Value}")
print(f"  B13 (Days Traded)   = {sh_sum.Range('B13').Value}")

# DON'T clear - let user see in Excel
# Re-protect
for sn in ["Summary","Daily Log","Raw daily data"]:
    s = wb.Sheets(sn)
    try: s.Unprotect(PWD)
    except: pass
    if sn == "Summary":
        for a in ["B2","B3","B5","B8","B9","B10"]: s.Range(a).Locked = False
    elif sn == "Daily Log":
        for r in ["A2:A100","B2:B100","D2:D100","M2:M100"]: s.Range(r).Locked = False
    elif sn == "Raw daily data":
        s.Range("A2:H500").Locked = False
    s.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
              AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
    s.EnableSelection = 0

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("\n=== MỞ FILE ĐỂ XEM KẾT QUẢ ===")
