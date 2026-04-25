"""Cure the formatting curse on Raw daily data."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass

# 1. Clear all weird formatting from row 2 downwards
sh.Range("A2:L500").ClearFormats()

# 2. Re-apply the basic font
sh.Range("A2:L500").Font.Name = "Calibri"
sh.Range("A2:L500").Font.Size = 11
sh.Range("A2:L500").Font.Color = 0

# 3. Re-apply the alternating background colors properly
for r in range(2, 501):
    if r % 2 == 0:
        sh.Range(f"A{r}:L{r}").Interior.Color = 16777215 # White
    else:
        sh.Range(f"A{r}:L{r}").Interior.Color = 15987699 # Very light blue

# 4. Re-apply the Profit formatting (Green for +, Red for -)
sh.Range("F2:F500").NumberFormat = "[Color10]$#,##0.00;[Red]-$#,##0.00;0"
sh.Range("H2:H500").NumberFormat = "m/d/yyyy h:mm"

# 5. Lock formula columns, unlock input columns
sh.Range("A2:L500").Locked = False
sh.Range("J2:L500").Locked = True # J, K, L are formulas (Cumulative P&L etc)

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

wb.Save(); wb.Close(); excel.Quit()
print("Cured Raw daily data formatting.")
