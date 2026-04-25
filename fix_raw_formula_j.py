"""Fix J formula to handle empty rows gracefully and format F."""
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

# 1. Format F as Currency and E as Number
sh.Range("F2:F500").NumberFormat = "$#,##0.00_ ;[Red]-$#,##0.00 "
sh.Range("E2:E500").NumberFormat = "#,##0.0000"

# 2. Fix J formula to not crash if Date (I) or previous J is blank
for r in range(2, 501):
    prev_j = "J1" if r == 2 else f"J{r-1}"
    prev_i = "I1" if r == 2 else f"I{r-1}"
    formula = f'=IF(OR(F{r}="", I{r}=""), "", F{r} + IF(I{r}={prev_i}, N({prev_j}), 0))'
    sh.Cells(r, 10).Formula = formula

excel.CalculateFullRebuild()

# Re-protect
sh.Range("A2:H500").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed format and #VALUE! in Raw daily data.")
