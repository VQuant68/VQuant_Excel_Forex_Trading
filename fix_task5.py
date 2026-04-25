"""
Explicitly apply all formula fixes from Task 5 to ensure 100% compliance.
"""
import win32com.client
import os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

print("=== APPLYING EXPLICIT FIXES FOR TASK 5 ===")
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
    
    sh_sum = wb.Sheets("Summary")
    sh_raw = wb.Sheets("Raw daily data")
    
    # 1. Summary B7
    print("Updating Summary!B7 with IFERROR...")
    sh_sum.Range("B7").Formula = '=IFERROR(B6/B17,"")'
    
    # 2. Summary B18
    print("Updating Summary!B18 with IFERROR...")
    sh_sum.Range("B18").Formula = '=IFERROR(IF(K22>0, B11/K22, ""),"")'
    
    # 3. Raw daily data column I (Trade Date)
    print("Updating Raw daily data!I2:I500 with IFERROR and Check date format...")
    # The formula needs to increment the row number: =IFERROR(IF(H2="","",DATEVALUE(MID(H2,6,11))),"Check date format")
    sh_raw.Range("I2:I500").Formula = '=IFERROR(IF(H2="","",DATEVALUE(MID(H2,6,11))),"Check date format")'
    
    wb.Save()
    wb.Close()
    excel.Quit()
    print("All fixes applied successfully!")
    
except Exception as e:
    print(f"Exception: {e}")
    try:
        excel.Quit()
    except:
        pass
