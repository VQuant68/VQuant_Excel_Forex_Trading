"""Inject 4 trades into Raw daily data for testing."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass

trades = [
    ["EURUSD", "EURUSD", "Sell", 1.0, 1.1600, -100.00, "102", "Tue, 09 May 2026 10:00:00 GMT"],
    ["EURUSD", "EURUSD", "Sell", 1.0, 1.1620, -100.00, "103", "Tue, 09 May 2026 15:00:00 GMT"],
    ["EURUSD", "EURUSD", "Buy", 2.0, 1.1500, 800.00, "104", "Wed, 17 May 2026 09:00:00 GMT"],
    ["EURUSD", "EURUSD", "Buy", 1.0, 1.1550, 200.00, "105", "Thu, 18 May 2026 14:00:00 GMT"]
]

# Find next empty row (should be row 3)
next_row = 3

for i, t in enumerate(trades):
    for j, val in enumerate(t):
        sh.Cells(next_row + i, j + 1).Value = val

sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0

wb.Save(); wb.Close(); excel.Quit()
print("Injected 4 trades.")
