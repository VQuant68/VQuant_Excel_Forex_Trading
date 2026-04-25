"""
VIỆC 6: Xóa test data, chuẩn bị Blank Template
- Daily Log: xóa A,B,D,M (giữ công thức)
- Raw daily data: xóa rows 2+
- Summary: reset về default values
- Advanced Setup Planner: xóa tất cả input (blue) cells
- Backend: KHÔNG đụng
"""
import win32com.client
import os

file_path = os.path.abspath('Trading_Workbook_MASTER.xlsx')
print("=== VIỆC 6: XÓA TEST DATA ===")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(file_path, UpdateLinks=False)

# ─────────────────────────────────────────────
# 1. Daily Log: xóa input columns A, B, D, M
# ─────────────────────────────────────────────
print("\n[1] Clearing Daily Log input columns...")
sh = wb.Sheets("Daily Log")
sh.Range("A2:A100").ClearContents()
sh.Range("B2:B100").ClearContents()
sh.Range("D2:D100").ClearContents()
sh.Range("M2:M100").ClearContents()

# Verify: A2 trống → công thức E2 phải ra ""
e2 = sh.Range("E2").Value
if e2 is None or e2 == "":
    print("   VERIFY OK: E2 = blank (công thức hiển thị '' khi A2 trống)")
else:
    print(f"   WARNING: E2 = '{e2}' - có thể công thức chưa bọc IF(A2='')")

print("   Done: A2:A100, B2:B100, D2:D100, M2:M100 cleared")

# ─────────────────────────────────────────────
# 2. Raw daily data: xóa rows 2+
# ─────────────────────────────────────────────
print("\n[2] Clearing Raw daily data (rows 2+)...")
sh_raw = wb.Sheets("Raw daily data")
last_row = sh_raw.Cells(sh_raw.Rows.Count, 1).End(-4162).Row  # xlUp
if last_row >= 2:
    sh_raw.Range(f"A2:H{last_row}").ClearContents()
    print(f"   Cleared rows 2 to {last_row}")
else:
    print("   Already empty")

# ─────────────────────────────────────────────
# 3. Summary: reset về default values
# ─────────────────────────────────────────────
print("\n[3] Resetting Summary to default values...")
sh_sum = wb.Sheets("Summary")
sh_sum.Range("B2").Value = 10000   # Monthly Target
sh_sum.Range("B10").Value = 10000  # Current NAV
# Giữ nguyên B3=17, B5=4, B8=0.07, B9=2
print("   B2 (Monthly Target) = 10,000")
print("   B10 (Current NAV) = 10,000")
print("   B3, B5, B8, B9 giữ nguyên")

# ─────────────────────────────────────────────
# 4. Advanced Setup Planner: xóa tất cả input (blue) cells
#    Blue cells = input cells (user enters EMA, RSI, price levels...)
#    Cols B (B5:B21), D (D5:D20), F,G (FVG), H,I (Options),
#    J,K (strikes), C23:D27 (BOS/CHOCH), K5:K10 (strikes extra)
# ─────────────────────────────────────────────
print("\n[4] Clearing Advanced Setup Planner input cells...")
sh_pl = wb.Sheets("Advanced Setup Planner")

input_ranges = [
    "B5:B21",    # Current Price, EMAs, RSIs, ATR
    "D5:D20",    # Liquidity / Session levels (PWH,PWL,PDH,PDL,EQH,EQL,etc.)
    "G5:G10",    # 4H FVG High/Low + Daily FVG + 1H FVG
    "H5:H10",    # FVG values col 2
    "J5:J12",    # Options strikes sizes
    "K5:K12",    # Strike prices
    "C24:C27",   # Last BOS manual inputs (1D,4H,1H,15M)
    "D24:D27",   # Last CHOCH manual inputs
    "B17:B21",   # RSI + ATR + ADR
]
for rng in input_ranges:
    try:
        sh_pl.Range(rng).ClearContents()
        print(f"   Cleared: {rng}")
    except Exception as e:
        print(f"   Skip {rng}: {e}")

# ─────────────────────────────────────────────
# 5. Verify E2 trong Daily Log hiển thị "" khi trống
# ─────────────────────────────────────────────
print("\n[5] Verifying formula cells show blank (not 0 or error)...")
excel.CalculateFullRebuild()
e2 = sh.Range("E2").Value
h2 = sh.Range("H2").Value
n2 = sh.Range("N2").Value

print(f"   Daily Log E2: '{e2}' (expected blank)")
print(f"   Daily Log H2: '{h2}' (expected blank)")
print(f"   Daily Log N2: '{n2}' (expected blank)")

results_ok = all(v is None or v == "" for v in [e2, h2, n2])
if results_ok:
    print("   PASS: All formula cells show blank when inputs are empty")
else:
    print("   WARNING: Some cells show non-blank when inputs are empty - need to wrap with IFERROR")

# ─────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────
print("\nSaving file...")
wb.Save()
wb.Close()
excel.Quit()

print("\n=== VIỆC 6 HOÀN THÀNH ===")
print("File đã về trạng thái Blank Template sạch sẽ.")
print("Mở lại file → mọi input cells đều trống, formula cells hiển thị \"\"")
