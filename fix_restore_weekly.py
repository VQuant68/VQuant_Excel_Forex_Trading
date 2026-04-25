"""Restore the Weekly Table that was accidentally cleared."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# 1. Restore Headers
headers = ["Total Trades", "Total Wins", "Total Losses", "Week Net P&L", "Weekly Target", "Week Variance ($)", "Avg Win Rate %", "Avg Reward:Risk", "Total R"]
for i, h in enumerate(headers):
    cell = sh.Cells(21, i + 3) # C21 onwards
    cell.Value = h
    cell.Font.Bold = True
    cell.Font.Color = 16777215 # White
    cell.Interior.Color = 5976593 # Dark blue (#11325B)
    cell.HorizontalAlignment = -4108 # Center

# 2. Restore Formulas for rows 22-25
for r in range(22, 26):
    # C: Total Trades (Sum of Daily Log N)
    sh.Cells(r, 3).Formula = f"=SUMIFS('Daily Log'!N:N, 'Daily Log'!B:B, $A{r})"
    # D: Total Wins (Sum of Daily Log O)
    sh.Cells(r, 4).Formula = f"=SUMIFS('Daily Log'!O:O, 'Daily Log'!B:B, $A{r})"
    # E: Total Losses (Sum of Daily Log P)
    sh.Cells(r, 5).Formula = f"=SUMIFS('Daily Log'!P:P, 'Daily Log'!B:B, $A{r})"
    
    # F: Week Net P&L (Sum of Daily Log H)
    sh.Cells(r, 6).Formula = f"=SUMIFS('Daily Log'!H:H, 'Daily Log'!B:B, $A{r})"
    # G: Weekly Target
    sh.Cells(r, 7).Formula = "=$B$6"
    # H: Week Variance
    sh.Cells(r, 8).Formula = f"=F{r}-G{r}"
    
    # I: Avg Win Rate %
    sh.Cells(r, 9).Formula = f'=IFERROR(D{r}/C{r}, 0)'
    # J: Avg Reward:Risk
    sh.Cells(r, 10).Formula = f"=IFERROR(SUMIFS('Daily Log'!T:T, 'Daily Log'!B:B, $A{r})/B{r}, 0)"
    # K: Total R
    sh.Cells(r, 11).Formula = f"=SUMIFS('Daily Log'!U:U, 'Daily Log'!B:B, $A{r})"

# 3. Restore Formatting
# Number formats
sh.Range("F22:H25").NumberFormat = "$#,##0.00;[Red]($#,##0.00);\"-\""
sh.Range("I22:I25").NumberFormat = "0.0%"
sh.Range("J22:K25").NumberFormat = "0.0"
sh.Range("C22:E25").NumberFormat = "0"

# Alignment
sh.Range("C22:K25").HorizontalAlignment = -4108 # Center
sh.Range("C21:K25").Font.Name = "Calibri"
sh.Range("C21:K25").Font.Size = 11

# Optional: Add borders to make it look like a table again
sh.Range("A21:K25").Borders.LineStyle = 1 # xlContinuous
sh.Range("A21:K25").Borders.Color = 14277081 # Light gray

sh.Columns("C:K").AutoFit()

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)
wb.Save(); wb.Close(); excel.Quit()
print("Restored Weekly Table successfully.")
