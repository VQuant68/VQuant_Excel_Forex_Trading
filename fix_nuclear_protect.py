"""Nuclear protection script: lock all, unlock specific inputs, reprotect."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

inputs = {
    "Summary": ["B2","B3","B5","B8","B9","B10"],
    "Daily Log": ["A2:A100", "B2:B100", "D2:D100", "M2:M100"],
    "Raw daily data": ["A2:H500"],
    "Advanced Setup Planner": ["B5:B21", "C5:D21", "E5:F11", "E13:E15", "J14:J20", "C24:D27"]
}

for sn in ["Summary", "Daily Log", "Raw daily data", "Advanced Setup Planner", "Instructions"]:
    sh = wb.Sheets(sn)
    try: sh.Unprotect(PWD)
    except: pass
    
    # 1. Lock EVERYTHING first
    sh.Cells.Locked = True
    
    # 2. Unlock only the specified input ranges
    if sn in inputs:
        for rng in inputs[sn]:
            try:
                sh.Range(rng).Locked = False
            except:
                pass
                
    # 3. Protect the sheet with column resizing allowed
    # Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows)
    sh.Protect(PWD, False, True, True, False, False, True, True)
    sh.EnableSelection = 0

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("Perfect Protection Applied.")
