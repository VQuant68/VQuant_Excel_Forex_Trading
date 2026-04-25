"""
Fix Summary!B11: LOOKUP picking up header row instead of data
Change range from $I:$I to $I$2:$I$100 to exclude header.
Also fix B12, B15, B16, B17 cascade errors.
"""
import win32com.client, os, openpyxl

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

# Read current formulas
wb_r = openpyxl.load_workbook(file_path, data_only=False)
ws = wb_r['Summary']
print("B11:", ws['B11'].value)
print("B12:", ws['B12'].value)
print("B13:", ws['B13'].value)
print("B15:", ws['B15'].value)
print("B16:", ws['B16'].value)
print("B17:", ws['B17'].value)
wb_r.close()

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Summary")

# Fix B11: Exclude header row by using $I$2:$I$100 instead of $I:$I
# Also sum approach is more reliable: =IFERROR(SUM('Daily Log'!H2:H100),"")
print("\nFixing B11 (Cumulative P&L Month)...")
sh.Range("B11").Formula = "=IFERROR(SUM('Daily Log'!$H$2:$H$100),\"\")"

# Fix B12: Remaining Target = B2 - B11
print("Fixing B12 (Remaining Target)...")
sh.Range("B12").Formula = '=IFERROR(IF(B11="",B2,B2-B11),"")'

# Fix B15: Required P&L per Remaining Day
print("Fixing B15...")
sh.Range("B15").Formula = '=IFERROR(IF(B14>0,B12/B14,""),"")'

# Fix B16: Target R per Day
print("Fixing B16...")
sh.Range("B16").Formula = '=IFERROR(IF(B8>0,B15/(B10*B8),""),"")'

# Fix B17: Projected Month P&L
print("Fixing B17...")
sh.Range("B17").Formula = '=IFERROR(IF(B11="","",B11/MAX(1,B13)*B3),"")' 

excel.CalculateFullRebuild()

# Verify
b11 = sh.Range("B11").Value
b12 = sh.Range("B12").Value
b13 = sh.Range("B13").Value
b15 = sh.Range("B15").Value
b16 = sh.Range("B16").Value
b17 = sh.Range("B17").Value

print(f"\nAfter fix (all inputs empty, no trades):")
print(f"  B11 (Cumul P&L): '{b11}' (expected 0 or blank)")
print(f"  B12 (Remaining Target): '{b12}' (expected 10000 or blank)")
print(f"  B13 (Trading Days Used): '{b13}' (expected 0)")
print(f"  B15 (Req P&L/Day): '{b15}' (expected blank)")
print(f"  B16 (Target R/Day): '{b16}' (expected blank)")
print(f"  B17 (Projected P&L): '{b17}' (expected blank)")

wb.Save()
wb.Close()
excel.Quit()
print("\nDone!")
