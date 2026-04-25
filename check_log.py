"""
Check Daily Log formulas for columns N, O, S, U to see what they depend on.
"""
import openpyxl

wb = openpyxl.load_workbook('Trading_Workbook_MASTER.xlsx', data_only=False)
ws = wb['Daily Log']

print("=== Formulas in Daily Log (Row 3) ===")
cols = {
    'H': 'Net P&L',
    'N': '# Trades',
    'O': '# Wins',
    'R': 'Total Losses',
    'S': 'Win Rate',
    'T': 'Reward:Risk',
    'U': 'Daily R-multiple',
    'W': 'Intraday Net P&L'
}

for col, name in cols.items():
    addr = f"{col}3"
    print(f"{addr} ({name}): {ws[addr].value}")
