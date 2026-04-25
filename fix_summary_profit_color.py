"""Fix the color format for profit and variance columns."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# Apply conditional formatting-like number format to Week Net P&L (F) and Week Variance (H)
# [Color10] is Green, [Red] is Red
format_str = "[Color10]$#,##0.00;[Red]($#,##0.00);\"-\""

# First, reset font color to automatic so the NumberFormat colors take over
sh.Range("F22:H25").Font.ColorIndex = 0 # Automatic

# Net P&L and Variance
sh.Range("F22:F25").NumberFormat = format_str
sh.Range("H22:H25").NumberFormat = format_str

# Weekly Target (G) can stay black
sh.Range("G22:G25").NumberFormat = "$#,##0.00"
sh.Range("G22:G25").Font.ColorIndex = 0

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed red profit issue.")
