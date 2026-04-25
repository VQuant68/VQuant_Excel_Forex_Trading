"""Fix font color on Advanced Setup Planner."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass

ranges_to_fix = [
    "B5:B21", "C5:D21", "E5:F11", "E13:E15", "J14:J20", "C24:D27"
]

for rng in ranges_to_fix:
    try:
        cell_range = sh.Range(rng)
        cell_range.Font.Color = 0  # Black color
        cell_range.Font.Bold = True # Make it bold so it's very easy to read
    except:
        pass

# Re-protect
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed Advanced Setup Planner font colors.")
