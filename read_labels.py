"""
VIỆC 9: Add tooltips (Comments) to Advanced Setup Planner input cells.
"""
import win32com.client, os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')

# First, read labels to verify cell mapping
print("=== Verifying cell labels ===")
for r in range(5, 22):
    label = sh.Cells(r, 1).Value  # col A = labels
    print(f"  A{r}: '{label}' → B{r}")
print()
for r in range(5, 19):
    label_l = sh.Cells(r, 4).Value  # col D = price level labels (left of E)
    print(f"  D{r}: '{label_l}' → E{r}")

# Check weights/magnet area
print("\nEngine params:")
for r in range(13, 22):
    lbl = sh.Cells(r, 8).Value  # col H area
    val = sh.Cells(r, 9).Value  # col I
    print(f"  H{r}: '{lbl}' = {val}")

# BOS/CHOCH
print("\nBOS/CHOCH:")
for r in range(23, 28):
    a = sh.Cells(r, 1).Value
    b = sh.Cells(r, 2).Value
    c = sh.Cells(r, 3).Value
    d = sh.Cells(r, 4).Value
    print(f"  Row {r}: A='{a}' B='{b}' C='{c}' D='{d}'")

# Side/Mode
print(f"\nA31: '{sh.Cells(31,1).Value}' B31: '{sh.Cells(31,2).Value}'")
print(f"A33: '{sh.Cells(33,1).Value}' B33: '{sh.Cells(33,2).Value}'")

wb.Close(False)
excel.Quit()
print("\nDone reading labels!")
