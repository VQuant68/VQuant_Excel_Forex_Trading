"""Fix #DIV/0! errors in Summary sheet when template is blank."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass
sh = wb.Sheets("Summary")
try: sh.Unprotect(PWD)
except: pass

# Find all #DIV/0! cells
print("=== Finding #DIV/0! errors ===")
used = sh.UsedRange
errors = []
for row in range(1, used.Rows.Count + 1):
    for col in range(1, used.Columns.Count + 1):
        cell = used.Cells(row, col)
        try:
            txt = str(cell.Text)
            if "#" in txt and ("DIV" in txt or "REF" in txt or "VALUE" in txt or "N/A" in txt):
                addr = cell.Address.replace("$","")
                formula = cell.Formula
                errors.append((addr, txt, formula))
                print(f"  {addr}: {txt} | formula: {formula}")
        except:
            pass

# Fix each error by wrapping with IFERROR
print(f"\nFound {len(errors)} errors. Fixing...")
for addr, txt, formula in errors:
    cell = sh.Range(addr)
    old = cell.Formula
    if old.startswith("=") and "IFERROR" not in old:
        new = f'=IFERROR({old[1:]},"")'
        cell.Formula = new
        print(f"  {addr}: wrapped with IFERROR")
    elif "IFERROR" in old:
        print(f"  {addr}: already has IFERROR, skipping")

excel.CalculateFullRebuild()

# Verify
print("\n=== Verification ===")
for addr, _, _ in errors:
    txt = sh.Range(addr).Text
    print(f"  {addr}: '{txt}'")

# Re-protect
sh.Range("B2").Locked = False
sh.Range("B3").Locked = False
sh.Range("B5").Locked = False
sh.Range("B8").Locked = False
sh.Range("B9").Locked = False
sh.Range("B10").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("\n=== Summary #DIV/0! fixed! ===")
