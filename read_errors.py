"""Fix circular reference in U2 and verify all formulas."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Daily Log")
try: sh.Unprotect(PWD)
except: pass

# Trace the dependency chain for U2
print("=== Tracing U2 dependencies ===")
cols_to_check = {
    'E': 5, 'F': 6, 'G': 7, 'H': 8, 'I': 9, 'J': 10,
    'K': 11, 'L': 12, 'N': 14, 'O': 15, 'P': 16,
    'Q': 17, 'R': 18, 'S': 19, 'T': 20, 'U': 21,
    'V': 22, 'W': 23, 'X': 24, 'Y': 25
}
for letter, num in cols_to_check.items():
    f = sh.Cells(2, num).Formula
    print(f"  {letter}2: {f}")

# Check Summary formulas that might reference Daily Log
sh_sum = wb.Sheets("Summary")
try: sh_sum.Unprotect(PWD)
except: pass
print("\n=== Summary cells that might cause circular ref ===")
for addr in ["K22","B4","B6","B8","B11","B13"]:
    f = sh_sum.Range(addr).Formula
    print(f"  Summary!{addr}: {f}")

# Check if I2 has cumulative reference to itself
print(f"\n=== I2 (Cumul P&L) ===")
print(f"  I2: {sh.Cells(2, 9).Formula}")
print(f"  I3: {sh.Cells(3, 9).Formula}")

wb.Close(False); excel.Quit()
