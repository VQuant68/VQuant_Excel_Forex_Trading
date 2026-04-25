"""
Fix E2 (hardcoded NAV value) and fix E column to show blank when A is empty.
"""
import win32com.client, os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Daily Log")

# E2 should reference Summary!K22 (Starting NAV) when A2 has a date
# or blank when A2 is empty
# E3+ should be =IF($A3="","",E2+H2)

# Fix E2: =IF($A2="","",Summary!K22)  (starting NAV from Summary)
sh.Range("E2").Formula = '=IF($A2="","",Summary!K22)'
print("E2 fixed: =IF($A2='','',Summary!K22)")

# Fix E3:E100 pattern: =IF($A3="","",E2+H2)
for row in range(3, 101):
    prev = row - 1
    sh.Cells(row, 5).Formula = f'=IF($A{row}="","",E{prev}+H{prev})'

print("E3:E100 fixed: =IF($Ax='','',Ex-1+Hx-1)")

excel.CalculateFullRebuild()
e2 = sh.Range("E2").Value
print(f"\nE2 after fix (A2 empty): '{e2}' (expected blank)")

wb.Save()
wb.Close()
excel.Quit()
print("Done! Việc 6 HOÀN THÀNH 100%")
