"""
Visual test: Force each value into B31/B33 directly.
Excel opens VISIBLE, pauses 5s per test for user to see.
File NOT saved.
"""
import win32com.client, os, time

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')
sh.Activate()
sh.Range('A29').Select()

tests = [
    ("TEST 1: LONG + CoreLong  → cả hai XANH LÁ", "LONG", "CoreLong"),
    ("TEST 2: LONG + A+Long   → cả hai XANH LÁ", "LONG", "A+Long"),
    ("TEST 3: SHORT + CTshort → cả hai ĐỎ NHẠT", "SHORT", "CTshort"),
    ("TEST 4: (trống) + Avoid → B33 ĐỎ ĐẬM", "", "Avoid"),
]

for label, side, mode in tests:
    print(f"\n{'='*55}")
    print(f"  {label}")
    print(f"  B31 = '{side}' | B33 = '{mode}'")
    print(f"{'='*55}")
    
    sh.Range('B31').Value = side
    sh.Range('B33').Value = mode
    
    time.sleep(5)  # User has 5 seconds to look

print("\n\nAll tests done! Look at Excel now.")
print("File will NOT be saved.")

# Wait a bit more for user to see last test
time.sleep(3)

wb.Close(False)
excel.Quit()
print("Excel closed (no changes saved).")
