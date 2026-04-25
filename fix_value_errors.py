"""Fix L2 #VALUE! and Y2 #DIV/0! in Daily Log."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass
sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass

# Fix L (col 12): Variance = J - K, but K might be "" 
# Fix Y (col 25): Drawdown divides by G which might be 0/""
# Fix V (col 22): also divides by G
# Fix X (col 24): also depends on G

fixes = {
    12: '=IF($A{r}="","",IFERROR(J{r}-K{r},""))',       # L = Variance
    25: '=IF($A{r}="","",IFERROR(ABS(MIN(0,MINIFS(\'Raw daily data\'!$J:$J,\'Raw daily data\'!$I:$I,$A{r})))/G{r},""))',  # Y
    22: '=IF($A{r}="","",IFERROR(IF(H{r}<0,MIN(ABS(H{r})/G{r},1),0),""))',  # V
    24: '=IF($A{r}="","",IFERROR(IF(W{r}>=0,0,ABS(W{r})/G{r}),""))',  # X
}

for col_num, template in fixes.items():
    col_letter = chr(64 + col_num) if col_num <= 26 else chr(64 + col_num - 26)
    for row in range(2, 42):
        formula = template.replace("{r}", str(row))
        sh.Cells(row, col_num).Formula = formula
    print(f"  Fixed col {col_num} (rows 2-41)")

excel.CalculateFullRebuild()

# Verify row 2
print("\n=== Verification (row 2, blank template) ===")
for col_num in [12, 22, 24, 25]:
    txt = sh.Cells(2, col_num).Text
    print(f"  Col {col_num}: '{txt}'")

# Re-protect
for rng in ["A2:A100","B2:B100","D2:D100","M2:M100"]:
    sh.Range(rng).Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("\n=== Fixed! ===")
