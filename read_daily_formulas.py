"""
Fix #VALUE! errors in Daily Log - wrap formulas with IF(A2="","",...)
"""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets('Daily Log')
try: sh.Unprotect(PWD)
except: pass

# First, read current formulas in row 2 for all formula columns
print("=== Current formulas in row 2 ===")
for col_letter, col_num in [
    ('E',5), ('F',6), ('G',7), ('H',8), ('I',9), ('J',10),
    ('K',11), ('L',12), ('N',14), ('O',15), ('P',16),
    ('Q',17), ('R',18), ('S',19), ('T',20), ('U',21),
    ('V',22), ('W',23), ('X',24), ('Y',25)
]:
    cell = sh.Cells(2, col_num)
    formula = cell.Formula
    value = cell.Value
    has_error = cell.Errors is not None
    print(f"  {col_letter}2: formula='{formula}' | value={value}")

wb.Close(False); excel.Quit()
print("\nDone reading!")
