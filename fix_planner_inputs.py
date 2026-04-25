"""
Fix Advanced Setup Planner:
Clear col E price levels cell by cell (avoid merged cell errors)
"""
import win32com.client, os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Advanced Setup Planner")

# Clear E col cell by cell, skipping merged cells
print("Clearing col E price levels (cell by cell)...")
cleared = 0
skipped = 0
for row in range(5, 31):
    try:
        cell = sh.Cells(row, 5)  # col E
        val = cell.Value
        if val is not None and val != "":
            cell.Value = None
            cleared += 1
    except Exception as e:
        skipped += 1
        
print(f"   Cleared: {cleared}, Skipped (merged): {skipped}")

# Also clear col K (strike prices, if any)
print("Clearing additional input cols (K5:K12)...")
for row in range(5, 13):
    try:
        sh.Cells(row, 11).Value = None  # col K
    except:
        pass

excel.CalculateFullRebuild()

e5 = sh.Cells(5, 5).Value
e10 = sh.Cells(10, 5).Value
print(f"\nE5: '{e5}' (expected blank)")
print(f"E10: '{e10}' (expected blank)")

wb.Save()
wb.Close()
excel.Quit()
print("Done!")
