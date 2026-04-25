"""Fix formats and formulas in Raw daily data."""
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

# 1. Fix Format for Profit (F) and Price (E)
sh.Range("F2:F500").NumberFormat = "$#,##0.00_ ;[Red]-$#,##0.00 "
sh.Range("E2:E500").NumberFormat = "#,##0.0000"

# 2. Fix formulas in I and J to avoid #VALUE! when empty
for r in range(2, 501):
    # Trade Date (I)
    # Parse date from string like "Wed, 23 Apr 2025 09:00:00 GMT"
    # If H is empty, return ""
    old_I = sh.Cells(r, 9).Formula
    if old_I.startswith("="):
        # We wrap it in IFERROR and check for blank
        sh.Cells(r, 9).Formula = f'=IF(H{r}="","",IFERROR({old_I[1:]},""))'

    # Cumulative P&L (J)
    # Usually it's something like =IF(I2="","",SUMIF($I$2:I2,I2,$F$2:F2))
    old_J = sh.Cells(r, 10).Formula
    if old_J.startswith("="):
        sh.Cells(r, 10).Formula = f'=IF(I{r}="","",IFERROR({old_J[1:]},""))'

excel.CalculateFullRebuild()

# Re-protect
sh.Range("A2:H500").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed format and formulas in Raw daily data.")
