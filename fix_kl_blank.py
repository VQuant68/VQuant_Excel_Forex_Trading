"""
Fix Daily Log columns K and L to show blank when date (col A) is empty.
"""
import win32com.client, os, openpyxl

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

# Read current formulas
wb_r = openpyxl.load_workbook(file_path, data_only=False)
ws = wb_r['Daily Log']
print("K2:", ws['K2'].value)
print("L2:", ws['L2'].value)
print("J2:", ws['J2'].value)
wb_r.close()

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Daily Log")

fixed_k = 0
fixed_l = 0

for row in range(2, 101):
    # Fix K column (Weekly Target)
    k_cell = sh.Cells(row, 11)
    k_formula = k_cell.Formula
    if k_formula and "IF($A" not in k_formula and "IF(A" not in k_formula:
        inner = k_formula.lstrip("=")
        k_cell.Formula = f'=IF($A{row}="","",{inner})'
        fixed_k += 1
    
    # Fix L column (Week Variance)
    l_cell = sh.Cells(row, 12)
    l_formula = l_cell.Formula
    if l_formula and "IF($A" not in l_formula and "IF(A" not in l_formula:
        inner = l_formula.lstrip("=")
        l_cell.Formula = f'=IF($A{row}="","",{inner})'
        fixed_l += 1

print(f"Fixed K: {fixed_k} rows, L: {fixed_l} rows")

excel.CalculateFullRebuild()
k2 = sh.Range("K2").Value
l2 = sh.Range("L2").Value
print(f"\nAfter fix (A2 empty):")
print(f"  K2: '{k2}' (expected blank)")
print(f"  L2: '{l2}' (expected blank)")

wb.Save()
wb.Close()
excel.Quit()
print("Done!")
