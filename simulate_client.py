"""Simulate 3 days of trading for realistic client testing."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

for sn in ["Summary", "Daily Log", "Raw daily data"]:
    try: wb.Sheets(sn).Unprotect(PWD)
    except: pass

# 1. Summary Setup
sh_sum = wb.Sheets("Summary")
sh_sum.Range("B2").Value = 2000
sh_sum.Range("B3").Value = 20
sh_sum.Range("B5").Value = 4
sh_sum.Range("B8").Value = 0.05
sh_sum.Range("B9").Value = 2
sh_sum.Range("B10").Value = 10000

# 2. Raw Daily Data (3 days of trades)
sh_raw = wb.Sheets("Raw daily data")
trades = [
    # Day 1: +150
    ["EURUSD", "EURUSD", "Buy", 0.5, 1.1500, 100.0, "101", "Mon, 01 May 2026 09:00:00 GMT"],
    ["EURUSD", "EURUSD", "Buy", 0.5, 1.1520, 50.0, "102", "Mon, 01 May 2026 14:00:00 GMT"],
    # Day 2: -50
    ["EURUSD", "EURUSD", "Sell", 0.5, 1.1600, -20.0, "103", "Tue, 02 May 2026 10:00:00 GMT"],
    ["EURUSD", "EURUSD", "Buy", 0.5, 1.1580, -30.0, "104", "Tue, 02 May 2026 15:00:00 GMT"],
    # Day 3: +80
    ["EURUSD", "EURUSD", "Buy", 0.5, 1.1550, 80.0, "105", "Wed, 03 May 2026 11:00:00 GMT"],
]
for i, t in enumerate(trades):
    for j, val in enumerate(t):
        sh_raw.Cells(i+2, j+1).Value = val

# 3. Daily Log (Log the 3 days)
sh_log = wb.Sheets("Daily Log")
logs = [
    ["5/1/2026", 1, "Y"],
    ["5/2/2026", 1, "Y"],
    ["5/3/2026", 1, "Y"]
]
for i, log in enumerate(logs):
    sh_log.Cells(i+2, 1).Value = log[0] # Date
    sh_log.Cells(i+2, 2).Value = log[1] # Week
    sh_log.Cells(i+2, 4).Value = log[2] # Traded

excel.CalculateFullRebuild()

# Reprotect
for sn in ["Summary", "Daily Log", "Raw daily data"]:
    sh = wb.Sheets(sn)
    sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("Simulated 3 days of trading data successfully.")
