"""Fix Raw daily data formatting - reset all colors properly."""
import win32com.client, os

def xlc(h):
    h=h.lstrip('#'); return int(h[0:2],16)+int(h[2:4],16)*256+int(h[4:6],16)*65536

PWD = "cuongdeptrai"
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
try: wb.Unprotect(PWD)
except: pass

sh = wb.Sheets("Raw daily data")
try: sh.Unprotect(PWD)
except: pass

# Reset ALL cells in data area
print("Resetting all formatting A2:J500...")
rng = sh.Range("A2:J500")
rng.Interior.Color = xlc("#FFFFFF")     # White background
rng.Font.Color = xlc("#333333")         # Dark text
rng.Font.Name = "Calibri"
rng.Font.Size = 11

# Alternating rows (subtle stripe)
print("Applying alternating rows...")
for row in range(2, 101):
    if row % 2 == 1:  # Odd rows = light stripe
        sh.Range(f"A{row}:J{row}").Interior.Color = xlc("#F5F8FC")

# Light borders
print("Adding borders...")
BORDER_COLOR = xlc("#D8D8D8")
rng2 = sh.Range("A2:J100")
for edge in [7, 8, 9, 10, 11, 12]:
    try:
        b = rng2.Borders(edge)
        b.LineStyle = 1
        b.Weight = 1
        b.Color = BORDER_COLOR
    except:
        pass

# Header styling
hdr = sh.Range("A1:J1")
hdr.Interior.Color = xlc("#1A3A5C")
hdr.Font.Color = xlc("#FFFFFF")
hdr.Font.Bold = True
hdr.Font.Size = 11

# Column widths
sh.Columns("A").ColumnWidth = 14
sh.Columns("B").ColumnWidth = 12
sh.Columns("C").ColumnWidth = 8
sh.Columns("D").ColumnWidth = 10
sh.Columns("E").ColumnWidth = 10
sh.Columns("F").ColumnWidth = 10
sh.Columns("G").ColumnWidth = 10
sh.Columns("H").ColumnWidth = 36
sh.Columns("I").ColumnWidth = 14
sh.Columns("J").ColumnWidth = 24

# Re-protect
sh.Range("A2:H500").Locked = False
sh.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True,
           AllowFormattingColumns=True, AllowFormattingRows=True, AllowFiltering=True)
sh.EnableSelection = 0
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("Done! White bg + dark text + alternating stripes.")
