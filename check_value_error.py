"""
Check what is causing #VALUE! in E, F, G, I columns of Daily Log
"""
import openpyxl

wb = openpyxl.load_workbook('Trading_Workbook_MASTER.xlsx', data_only=False)
ws = wb['Daily Log']

print("=== Formulas in Daily Log (Row 5 and 6) ===")
cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']

for row in [5, 6]:
    print(f"Row {row}:")
    for col in cols:
        addr = f"{col}{row}"
        print(f"  {addr}: {ws[addr].value}")
