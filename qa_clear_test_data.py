"""Clear all test data to leave a clean blank template."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

for sn in ["Summary", "Daily Log", "Raw daily data", "Advanced Setup Planner"]:
    sh = wb.Sheets(sn)
    try: sh.Unprotect(PWD)
    except: pass

# Clear Summary
sh_sum = wb.Sheets("Summary")
for addr in ["B2","B3","B5","B8","B9","B10"]:
    sh_sum.Range(addr).Value = None

# Clear Daily Log
sh_dl = wb.Sheets("Daily Log")
for r in range(2, 42):
    for c in [1, 2, 4, 13]: # A, B, D, M
        sh_dl.Cells(r, c).Value = None

# Clear Raw daily data
sh_raw = wb.Sheets("Raw daily data")
for r in range(2, 100):
    for c in range(1, 9):
        sh_raw.Cells(r, c).Value = None

# Clear Planner
sh_plan = wb.Sheets("Advanced Setup Planner")
for r in range(5, 22):
    sh_plan.Cells(r, 2).Value = None
sh_plan.Cells(13, 5).Value = None
sh_plan.Cells(14, 5).Value = None

excel.CalculateFullRebuild()

# Reprotect
for sn in ["Summary", "Daily Log", "Raw daily data", "Advanced Setup Planner"]:
    sh = wb.Sheets(sn)
    sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
               AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
    sh.EnableSelection = 0

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("All test data cleared! File is now a clean template.")
