"""Hide all comments - hover only."""
import win32com.client, os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')

excel.DisplayCommentIndicator = 1

for i in range(1, sh.Comments.Count + 1):
    sh.Comments(i).Visible = False

print(f"Hidden {sh.Comments.Count} comments.")
wb.Save(); wb.Close(); excel.Quit()
print("Done!")
