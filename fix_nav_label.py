"""Fix NAV label and add Current NAV display."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# 1. Rename A10 to Starting NAV
sh.Range("A10").Value = "Starting NAV ($)"

# 2. Add Current NAV to D10 and E10 so they can see it
sh.Range("C10").Value = "Current NAV ($) ➔"
sh.Range("C10").Font.Bold = True
sh.Range("C10").HorizontalAlignment = -4152 # Right align

sh.Range("D10").Formula = "=B10+B11"
sh.Range("D10").Font.Bold = True
sh.Range("D10").Font.Color = 0
sh.Range("D10").Interior.Color = 13434828 # Light green to highlight
sh.Range("D10").NumberFormat = "$#,##0.00"

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)
wb.Save(); wb.Close(); excel.Quit()
print("Fixed NAV labels on Summary.")
