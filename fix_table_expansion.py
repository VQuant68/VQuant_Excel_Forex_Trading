"""Fix Excel Table auto-expansion issue on Summary."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# 1. Clear Columns C, D, E, F completely
sh.Columns("C:Z").Clear()

# 2. Resize any tables (ListObjects) to only cover columns A and B
for obj in sh.ListObjects:
    # Get current range of the table
    rng = obj.Range
    # If it expanded past column B, resize it
    if rng.Columns.Count > 2:
        new_rng = sh.Range(sh.Cells(rng.Row, 1), sh.Cells(rng.Row + rng.Rows.Count - 1, 2))
        obj.Resize(new_rng)

# 3. Add Current NAV to E10 and F10 (Leave C and D completely blank so Table doesn't expand)
sh.Columns("E:E").ColumnWidth = 15
sh.Columns("F:F").ColumnWidth = 15

sh.Range("E10").Value = "Current NAV ➔"
sh.Range("E10").Font.Bold = True
sh.Range("E10").HorizontalAlignment = -4152 # xlRight

sh.Range("F10").Formula = "=B10+B11"
sh.Range("F10").Font.Bold = True
sh.Range("F10").Font.Color = 0
sh.Range("F10").Interior.Color = 13434828 # Light green
sh.Range("F10").NumberFormat = "$#,##0.00"

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)
wb.Save(); wb.Close(); excel.Quit()
print("Fixed table auto-expansion.")
