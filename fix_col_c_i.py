"""
Fix Daily Log:
  - Col C (Day of Week): IF(A="","",TEXT(A,"dddd"))
  - Col I (Cumulative P&L Month): wrap IFERROR
"""
import win32com.client, os, openpyxl

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

# First read current formulas
wb_r = openpyxl.load_workbook(file_path, data_only=False)
ws = wb_r['Daily Log']
print("Col C row 3:", ws['C3'].value)
print("Col I row 2:", ws['I2'].value)
print("Col I row 3:", ws['I3'].value)
wb_r.close()

# Fix via COM
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Daily Log")

print("\nFixing Col C (Day of Week)...")
for row in range(2, 101):
    sh.Cells(row, 3).Formula = f'=IF($A{row}="","",TEXT($A{row},"dddd"))'

print("Fixing Col I (Cumulative P&L Month)...")
# Read current I formulas first
for row in range(2, 101):
    cell = sh.Cells(row, 9)  # Col I
    formula = cell.Formula
    if formula and formula.startswith("="):
        # Wrap with IFERROR if not already
        if "IFERROR" not in formula:
            cell.Formula = f'=IFERROR({formula[1:]},"")' 
    elif not formula:
        # If blank/value, set to SUMIF pattern
        pass

excel.CalculateFullRebuild()

# Verify
c2 = sh.Range("C2").Value
i2 = sh.Range("I2").Value
i3 = sh.Range("I3").Value
print(f"\nAfter fix:")
print(f"  C2 (Day of Week, A2 empty): '{c2}' (expected blank)")
print(f"  I2 (Cumul P&L Month): '{i2}' (expected blank)")
print(f"  I3 (Cumul P&L Month): '{i3}' (expected blank)")

wb.Save()
wb.Close()
excel.Quit()
print("\nDone!")
