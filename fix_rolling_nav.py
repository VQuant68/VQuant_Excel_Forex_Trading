"""Fix Rolling NAV and Cumulative P&L in Daily Log."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh_dl = wb.Sheets("Daily Log")
try: sh_dl.Unprotect(PWD)
except: pass

# 1. Fix Start NAV (Column E)
sh_dl.Range("E2").Formula = "=Summary!$B$10"
sh_dl.Range("E3").Formula = '=IF($A3="","", E2+H2)'
sh_dl.Range("E3:E100").FillDown()

# 2. Fix Cumulative P&L (Column I)
sh_dl.Range("I2").Formula = '=IF($A2="","", H2)'
sh_dl.Range("I3").Formula = '=IF($A3="","", I2+H3)'
sh_dl.Range("I3:I100").FillDown()

sh_dl.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

# 3. Check Summary B11 formula just in case
sh_sum = wb.Sheets("Summary")
try: sh_sum.Unprotect(PWD)
except: pass
sh_sum.Range("B11").Formula = "=SUM('Daily Log'!$H$2:$H$100)"
sh_sum.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

# Force full calculation to ensure all numbers update
wb.Application.CalculateFull()

wb.Save(); wb.Close(); excel.Quit()
print("Fixed rolling NAV and Cumulative P&L formulas.")
