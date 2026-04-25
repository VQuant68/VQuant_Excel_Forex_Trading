"""
Apply Data Validations directly to MASTER workbook (Task 5)
"""
import win32com.client
import os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')

print("=== APPLYING DATA VALIDATION (TASK 5) ===")
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
    
    sh_log = wb.Sheets("Daily Log")
    sh_raw = wb.Sheets("Raw daily data")
    
    # Validation for Daily Log D2:D100 (Y,N)
    rng1 = sh_log.Range("D2:D100")
    rng1.Validation.Delete()
    rng1.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="Y,N")
    rng1.Validation.InCellDropdown = True
    print("Added Y,N validation to Daily Log D2:D100")

    # Validation for Raw daily data C2:C500 (buy,sell)
    rng2 = sh_raw.Range("C2:C500")
    rng2.Validation.Delete()
    rng2.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1="buy,sell")
    rng2.Validation.InCellDropdown = True
    print("Added buy,sell validation to Raw daily data C2:C500")

    wb.Save()
    wb.Close()
    excel.Quit()
    print("Validations saved successfully.")
    
except Exception as e:
    print(f"Exception: {e}")
    try:
        excel.Quit()
    except:
        pass
