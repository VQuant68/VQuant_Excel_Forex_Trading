"""
Full audit: check all input cells in Advanced Setup Planner should be blank.
"""
import win32com.client, os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(file_path, UpdateLinks=False)
sh = wb.Sheets("Advanced Setup Planner")

problems = []

# Input cells that should be blank (blue cells)
input_cells = {
    # Col B: Price and EMAs
    "B5": "Current Price",
    "B6": "EMA9 15m",
    "B7": "EMA21 15m",
    "B8": "EMA50 15m",
    "B9": "SMA200 15m",
    "B10": "EMA50 1H",
    "B11": "EMA50 4H",
    "B12": "EMA50 1D",
    "B13": "EMA100 1D",
    "B14": "EMA200 1D",
    "B15": "EMA50 1W",
    "B16": "EMA20 1M",
    "B17": "RSI 15m",
    "B18": "ATR 15m",
    "B19": "ADR",
    "B20": "RSI 1H",
    "B21": "RSI 4H",
    # Col E: Price Levels
    "E5": "PMH",
    "E6": "PML",
    "E7": "PWH",
    "E8": "PWL",
    "E9": "PDH",
    "E10": "PDL",
    "E11": "EQH_above",
    "E12": "EQL_below",
    "E13": "LTF Leg Low 15m",
    "E14": "LTF Leg High 15m",
    "E15": "Session Low",
    "E16": "Session High",
    "E17": "HTF Swing Low 4H",
    "E18": "HTF Swing High 4H",
    # BOS/CHOCH
    "C24": "BOS 1D",
    "C25": "BOS 4H",
    "C26": "BOS 1H",
    "C27": "BOS 15M",
    "D24": "CHOCH 1D",
    "D25": "CHOCH 4H",
    "D26": "CHOCH 1H",
    "D27": "CHOCH 15M",
}

print("=== AUDIT: Advanced Setup Planner Input Cells ===\n")
all_pass = True
for addr, label in input_cells.items():
    val = sh.Range(addr).Value
    if val is not None and val != "" and val != 0:
        print(f"  ❌ {addr} ({label}): '{val}' — NOT BLANK!")
        problems.append(addr)
        all_pass = False
    else:
        print(f"  ✅ {addr} ({label}): blank")

print(f"\n=== RESULT ===")
if all_pass:
    print("ALL INPUT CELLS ARE BLANK — Việc 6 HOÀN THÀNH 100%!")
else:
    print(f"CÒN {len(problems)} Ô CHƯA TRỐNG: {problems}")
    print("Clearing them now...")
    for addr in problems:
        sh.Range(addr).Value = None
    wb.Save()
    print("Fixed and saved!")

if not all_pass:
    wb.Save()
wb.Close()
excel.Quit()
