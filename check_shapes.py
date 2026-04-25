import win32com.client, os
PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')
try: sh.Unprotect(PWD)
except: pass

print(f"Comments: {sh.Comments.Count}")
print(f"Shapes: {sh.Shapes.Count}")
for i in range(1, sh.Shapes.Count+1):
    s = sh.Shapes(i)
    print(f"  Shape {i}: '{s.Name}' type={s.Type} pos=({s.Left:.0f},{s.Top:.0f})")

wb.Close(False); excel.Quit()
