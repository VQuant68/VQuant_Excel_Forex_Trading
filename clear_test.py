import win32com.client, os
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath("Trading_Workbook_MASTER.xlsx"))
sh = wb.Sheets("Advanced Setup Planner")
for r in range(5, 22):
    try: sh.Cells(r, 2).Value = None
    except: pass
excel.CalculateFullRebuild()
b31 = sh.Range("B31").Value
b33 = sh.Range("B33").Value
print(f"B31: '{b31}' | B33: '{b33}'")
wb.Save(); wb.Close(); excel.Quit()
print("Test data cleared!")
