"""
VIỆC 11 v2: Reposition navigation buttons to non-overlapping locations.
Delete old shapes, add new ones in empty areas.
"""
import win32com.client, os

PWD = "cuongdeptrai"

def xlc(h):
    h=h.lstrip('#'); return int(h[0:2],16)+int(h[2:4],16)*256+int(h[4:6],16)*65536

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))

try: wb.Unprotect(PWD)
except: pass

def delete_all_shapes(ws):
    """Delete all shapes (nav buttons) from sheet."""
    count = ws.Shapes.Count
    for i in range(count, 0, -1):
        try: ws.Shapes(i).Delete()
        except: pass
    return count

def add_nav(ws, left, top, label, target_sheet, target_cell="A1", 
            bg="#1A3A5C", w=150, h=26):
    shp = ws.Shapes.AddShape(5, left, top, w, h)  # RoundedRectangle
    shp.TextFrame2.TextRange.Text = label
    shp.TextFrame2.TextRange.Font.Size = 9.5
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = xlc("#FFFFFF")
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shp.TextFrame2.VerticalAnchor = 3
    shp.Fill.ForeColor.RGB = xlc(bg)
    shp.Line.Visible = False
    shp.Shadow.Visible = True
    shp.Shadow.Type = 21  # outer shadow
    shp.Shadow.Transparency = 0.7
    shp.Shadow.Size = 100
    shp.Shadow.Blur = 4
    shp.Shadow.OffsetX = 2
    shp.Shadow.OffsetY = 2
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress=f"'{target_sheet}'!{target_cell}")
    return shp

# Get Instructions section rows
inst = wb.Sheets("Instructions")
sec5_row, sec6_row = 1, 1
for r in range(1, 200):
    val = inst.Cells(r, 2).Value
    if val and "SECTION 5" in str(val): sec5_row = r
    if val and "SECTION 6" in str(val): sec6_row = r
print(f"Section 5 row: {sec5_row}, Section 6 row: {sec6_row}")

# ══════════════════════════════════════════════
# SUMMARY — button at very top-right, above data
# ══════════════════════════════════════════════
print("\n── Summary ──")
ws = wb.Sheets("Summary")
try: ws.Unprotect(PWD)
except: pass
deleted = delete_all_shapes(ws)
print(f"  Deleted {deleted} old shapes")

# Place at top-right corner of the visible area
# Row 1, after last data column (use pixel position)
left = ws.Cells(1, 7).Left  # column G area
top = 4  # top margin
add_nav(ws, left, top, "📖  Instructions", "Instructions", bg="#95A5A6", w=140)
print("  Added [📖 Instructions] at G1 area")

ws.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
ws.EnableSelection = 0

# ══════════════════════════════════════════════
# DAILY LOG — button at right end of header row
# ══════════════════════════════════════════════
print("\n── Daily Log ──")
ws = wb.Sheets("Daily Log")
try: ws.Unprotect(PWD)
except: pass
deleted = delete_all_shapes(ws)
print(f"  Deleted {deleted} old shapes")

# Place after last visible column (Y=25, so col 26+)
left = ws.Cells(1, 26).Left  # after col Y
top = 4
add_nav(ws, left, top, "←  Summary", "Summary", w=120)
print("  Added [← Summary] after col Y")

ws.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
ws.EnableSelection = 0

# ══════════════════════════════════════════════
# RAW DAILY DATA — buttons after col J
# ══════════════════════════════════════════════
print("\n── Raw daily data ──")
ws = wb.Sheets("Raw daily data")
try: ws.Unprotect(PWD)
except: pass
deleted = delete_all_shapes(ws)
print(f"  Deleted {deleted} old shapes")

left = ws.Cells(1, 12).Left  # col L area
top = 4
add_nav(ws, left, top, "←  Summary", "Summary", w=120)
add_nav(ws, left + 130, top, "📋  How to Import", "Instructions", f"A{sec5_row}", bg="#2E86AB", w=150)
print("  Added 2 buttons at L1 area")

ws.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
ws.EnableSelection = 0

# ══════════════════════════════════════════════
# ADVANCED SETUP PLANNER — buttons in row 1 empty area
# ══════════════════════════════════════════════
print("\n── Advanced Setup Planner ──")
ws = wb.Sheets("Advanced Setup Planner")
try: ws.Unprotect(PWD)
except: pass
deleted = delete_all_shapes(ws)
print(f"  Deleted {deleted} old shapes")

# Row 2 has instruction text. Put nav at col F-G area row 1
left = ws.Cells(1, 6).Left   # col F
top = 4
add_nav(ws, left, top, "←  Summary", "Summary", w=120)
add_nav(ws, left + 130, top, "📋  How to Use Planner", "Instructions", f"A{sec6_row}", bg="#E67E22", w=170)
print("  Added 2 buttons at F1 area")

ws.Protect(Password=PWD, DrawingObjects=False, Contents=True, Scenarios=True, AllowFiltering=True)
ws.EnableSelection = 0

# Re-protect workbook
wb.Protect(Password=PWD, Structure=True, Windows=False)

wb.Save(); wb.Close(); excel.Quit()
print("\n=== VIỆC 11 v2 HOÀN THÀNH ===")
