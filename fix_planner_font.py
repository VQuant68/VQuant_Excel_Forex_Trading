"""Fix font formatting on Advanced Setup Planner after copy-paste."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Advanced Setup Planner")
try: sh.Unprotect(PWD)
except: pass

# List of input ranges to enforce Calibri 11
ranges_to_fix = [
    "B5:B21",    # Market snapshot
    "C5:D21",    # Liquidity
    "E5:F11",    # FVG
    "E13:E15",   # Legs
    "J14:J20",   # Options expiry
    "C24:D27"    # BOS / CHOCH
]

for rng in ranges_to_fix:
    try:
        cell_range = sh.Range(rng)
        cell_range.Font.Name = "Calibri"
        cell_range.Font.Size = 11
        # Set text alignment to right for number columns to look clean
        if rng in ["B5:B21", "E13:E15"]:
            cell_range.HorizontalAlignment = -4152 # xlRight
    except:
        pass

# Re-protect
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Fixed Advanced Setup Planner fonts.")
