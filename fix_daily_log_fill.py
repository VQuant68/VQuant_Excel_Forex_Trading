"""Fix Daily Log formatting and formulas."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass

# Copy row 2 to rows 3:100
# We will use FillDown to ensure formulas and formatting are copied perfectly
sh.Range("A2:W100").FillDown()

# Now clear the INPUT cells for rows 3:100 so the user can type in them
# Input columns are A (1), B (2), D (4), M (13)
sh.Range("A3:B100").ClearContents()
sh.Range("D3:D100").ClearContents()
sh.Range("M3:M100").ClearContents()

# Ensure the font is consistent
sh.Range("A2:W100").Font.Name = "Calibri"
sh.Range("A2:W100").Font.Size = 11

sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed Daily Log formatting and formulas.")
