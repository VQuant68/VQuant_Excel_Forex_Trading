import win32com.client, os
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')

cells_to_check = ['B5','B6','B7','B10','B11','B17','B20','E7','E9','C24','D24']
for addr in cells_to_check:
    cell = sh.Range(addr)
    try:
        cmt = cell.Comment.Text()
        print(f'{addr}: "{cmt[:70]}..."')
    except:
        print(f'{addr}: NO COMMENT')

total = sh.Comments.Count
print(f'\nTotal comments on sheet: {total}')
print(f'DisplayCommentIndicator: {excel.DisplayCommentIndicator}')

# Force indicator visible (1 = xlCommentIndicatorOnly)
excel.DisplayCommentIndicator = 1
print(f'Set to: {excel.DisplayCommentIndicator}')

wb.Save(); wb.Close(); excel.Quit()
print("Done!")
