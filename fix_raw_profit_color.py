"""Fix color formatting for Profit column in Raw daily data."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass

# Use Color10 (dark green) for positive, Red for negative
# Format: [Color10]$#,##0.00_ ;[Red]-$#,##0.00 ;[Black]$0.00_ 
sh.Range("F2:F500").NumberFormat = "[Color10]$#,##0.00_ ;[Red]-$#,##0.00 ;[Black]$0.00_ "

# Re-protect
sh.Range("A2:H500").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Profit column color format updated to Green/Red.")
