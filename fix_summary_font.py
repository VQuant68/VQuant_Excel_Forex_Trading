"""Fix font formatting on Summary sheet after copy-paste."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# Reset font for input cells
input_cells = ["B2","B3","B5","B8","B9","B10"]
for addr in input_cells:
    cell = sh.Range(addr)
    cell.Font.Name = "Calibri"
    cell.Font.Size = 11
    cell.Font.Color = 0  # Black
    
# Fix Row 1 if it got messed up (clear anything from C1:Z1)
try:
    sh.Range("C1:Z1").ClearContents()
except:
    pass

# Re-protect
for addr in input_cells:
    sh.Range(addr).Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("Fixed Summary fonts and cleaned up Row 1.")
