"""Clean up garbage data in Summary sheet rows 1-20."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# Clear garbage data in rows 1 to 20, columns C to Z
# But be careful not to touch the Weekly Table which starts at row 21
sh.Range("C1:Z9").Clear()
sh.Range("E10:Z10").Clear()
sh.Range("C11:Z20").Clear()

# Re-apply the Current NAV in C10 and D10 properly
sh.Range("C10").Value = "Current NAV ➔"
sh.Range("C10").Font.Bold = True
sh.Range("C10").HorizontalAlignment = -4152 # xlRight

sh.Range("D10").Formula = "=B10+B11"
sh.Range("D10").Font.Bold = True
sh.Range("D10").Font.Color = 0
sh.Range("D10").Interior.Color = 13434828 # Light green
sh.Range("D10").NumberFormat = "$#,##0.00"

# Reprotect
sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)
wb.Save(); wb.Close(); excel.Quit()
print("Cleared garbage data from Summary sheet.")
