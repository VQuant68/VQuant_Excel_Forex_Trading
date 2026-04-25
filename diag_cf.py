"""Diagnose conditional formatting on B31 and B33"""
import win32com.client, os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')

# Check actual values
b31 = sh.Range('B31')
b33 = sh.Range('B33')

print(f"B31 value: '{b31.Value}'")
print(f"B31 formula: '{b31.Formula}'")
print(f"B31 MergeArea: {b31.MergeArea.Address}")
print(f"B31 CF count: {b31.FormatConditions.Count}")
for i in range(1, b31.FormatConditions.Count+1):
    fc = b31.FormatConditions(i)
    print(f"  Rule {i}: Type={fc.Type}, Formula1={getattr(fc,'Formula1','N/A')}")
    print(f"  Interior.Color={fc.Interior.Color}")

print()
print(f"B33 value: '{b33.Value}'")
print(f"B33 formula: '{b33.Formula}'")
print(f"B33 CF count: {b33.FormatConditions.Count}")
for i in range(1, b33.FormatConditions.Count+1):
    fc = b33.FormatConditions(i)
    print(f"  Rule {i}: Type={fc.Type}, Formula1={getattr(fc,'Formula1','N/A')}")
    print(f"  Interior.Color={fc.Interior.Color}")

# Check if B31 is merged - get the top-left cell of merge area
merge_addr = b31.MergeArea.Address
print(f"\nB31 merge area: {merge_addr}")
print(f"B31 merge area rows: {b31.MergeArea.Rows.Count}, cols: {b31.MergeArea.Columns.Count}")

wb.Close(False)
excel.Quit()
