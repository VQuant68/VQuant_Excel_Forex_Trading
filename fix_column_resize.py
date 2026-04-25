"""Fix column resizing issue by applying correct positional arguments to Protect."""
import win32com.client, os

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

for sn in ["Advanced Setup Planner", "Raw daily data", "Daily Log", "Summary"]:
    sh = wb.Sheets(sn)
    try: sh.Unprotect(PWD)
    except: pass
    
    # AutoFit columns that might be too narrow in Planner
    if sn == "Advanced Setup Planner":
        sh.Columns("K:K").AutoFit()
        sh.Columns("C:C").AutoFit()
        sh.Columns("E:E").AutoFit()

    # Protect using positional arguments:
    # 1:Password, 2:DrawingObjects, 3:Contents, 4:Scenarios, 5:UserInterfaceOnly
    # 6:AllowFormattingCells, 7:AllowFormattingColumns, 8:AllowFormattingRows
    # 9:AllowInsertingColumns, 10:AllowInsertingRows, 11:AllowInsertingHyperlinks
    # 12:AllowDeletingColumns, 13:AllowDeletingRows, 14:AllowSorting, 15:AllowFiltering
    sh.Protect(PWD, False, True, True, False, False, True, True, False, False, False, False, False, False, True)
    sh.EnableSelection = 0

wb.Protect(Password=PWD, Structure=True, Windows=False)
wb.Save(); wb.Close(); excel.Quit()
print("Fixed protection to allow column resizing.")
