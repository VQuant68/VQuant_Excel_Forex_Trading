"""
VIỆC 9: Add Comments/Tooltips to Advanced Setup Planner input cells.
"""
import win32com.client, os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.abspath('Trading_Workbook_MASTER.xlsx'))
sh = wb.Sheets('Advanced Setup Planner')

# ── Tooltip mapping: (cell_address, tooltip_text) ────
tooltips = [
    # Col B: EMAs and indicators
    ("B5",  "Current live mid-price of EURUSD. Drives all trend calculations."),
    ("B6",  "EMA 9 from 15-minute chart. Short-term momentum. If price > EMA9, bullish pressure."),
    ("B7",  "EMA 21 from 15-min chart. Cross with EMA9 = momentum shift signal."),
    ("B8",  "EMA 50 on 15min. Key short-term structure level."),
    ("B9",  "200 SMA on 15min = major intraday support or resistance zone."),
    ("B10", "1H EMA50 = primary intraday trend filter. Price above = bullish context."),
    ("B11", "4H EMA50 = medium-term bias. Core filter for LONG vs SHORT decision."),
    ("B12", "Daily EMA50 = long-term trend anchor."),
    ("B13", "Daily EMA100 = intermediate trend level between EMA50 and EMA200."),
    ("B14", "Daily EMA200 = The golden line. Above = long-term bull regime."),
    ("B15", "Weekly EMA50 = major trend filter. Changes slowly."),
    ("B16", "Monthly EMA20 = macro direction anchor for swing context."),
    ("B17", "RSI(14) on 15min chart. Below 45 = bearish momentum. Above 55 = bullish."),
    ("B18", "RSI(14) on 1H chart. Below 45 supports LONG bias in scoring engine."),
    ("B19", "RSI(14) on 4H chart. HTF momentum confirmation."),
    ("B20", "Average True Range on 15min. Auto-sets stop floor in pips."),
    ("B21", "Average Daily Range in pips. Helps gauge if TP target is realistic today."),
    
    # Col E: Price levels (check label from col D)
    ("E5",  "Previous Month High. Major monthly liquidity level."),
    ("E6",  "Previous Month Low. Major monthly support level."),
    ("E7",  "Previous Week High. Key liquidity level — price often sweeps these."),
    ("E8",  "Previous Week Low. Key liquidity level — price often sweeps these."),
    ("E9",  "Previous Day High. Nearest intraday resistance."),
    ("E10", "Previous Day Low. Nearest intraday support."),
    ("E11", "Equal Highs above market = liquidity pool. May be targeted by stop hunt."),
    ("E12", "Equal Lows below market = buy-side liquidity below price."),
    ("E13", "15m Swing Leg Low — lower bound of most recent impulse move."),
    ("E14", "15m Swing Leg High — upper bound of most recent impulse move."),
    ("E15", "Current session low — intraday floor."),
    ("E16", "Current session high — intraday ceiling."),
    ("E17", "4H HTF Swing Low — higher timeframe support structure."),
    ("E18", "4H HTF Swing High — higher timeframe resistance structure."),
    
    # BOS/CHOCH inputs (C24:D27)
    ("C24", "Break of Structure on 1D — most recent confirmed BOS direction (from your chart)."),
    ("C25", "Break of Structure on 4H — most recent confirmed BOS direction."),
    ("C26", "Break of Structure on 1H — most recent confirmed BOS direction."),
    ("C27", "Break of Structure on 15M — most recent confirmed BOS direction."),
    ("D24", "Change of Character on 1D — signals potential trend reversal. Observe manually."),
    ("D25", "Change of Character on 4H — signals potential trend reversal."),
    ("D26", "Change of Character on 1H — signals potential trend reversal."),
    ("D27", "Change of Character on 15M — signals potential trend reversal."),
    
    # Outputs (read-only info)
    ("B31", "AUTO. Do not edit. Determined by HTF EMA alignment across 1D + 4H."),
    ("B33", "AUTO. CoreLong=follow trend; A+Long=premium entry; CTshort=fade; Avoid=no trade."),
]

# ── Add engine param tooltips (find them dynamically) ──
# Search for Magnet Range, Weights, Min R in the sheet
for row in range(1, 50):
    for col in range(1, 20):
        val = sh.Cells(row, col).Value
        if val is None: continue
        val = str(val).strip()
        addr = sh.Cells(row, col+1).Address.replace("$","")
        
        if "Magnet R" in val and "(" not in val:
            tooltips.append((addr, "Distance in pips within which a strike is 'active'. Default: 20."))
        elif "Options Clamp" in val or "Options Cl" in val:
            tooltips.append((addr, "Options gamma clamp percentage. Controls how much gamma influences price."))
        elif "Zone tolerance" in val or "Zone tol" in val:
            tooltips.append((addr, "Tolerance zone for price proximity to strike levels."))
        elif val == "Weight 0.382" or "0.382" in val and "Weight" in val:
            tooltips.append((addr, "Position size weight for Fib 0.382 entry rung. Default: 0.20. Must sum to 1.0."))
        elif val == "Weight 0.500" or "0.500" in val and "Weight" in val:
            tooltips.append((addr, "Position size weight for Fib 0.500 entry rung. Default: 0.35."))
        elif val == "Weight 0.618" or "0.618" in val and "Weight" in val:
            tooltips.append((addr, "Position size weight for Fib 0.618 entry rung. Default: 0.45."))
        elif val == "Min R (R multiple)" or "Min R" in val:
            tooltips.append((addr, "Minimum R-multiple to accept a trade. Below this threshold = skip. Default: 3."))
        elif val == "Ticket Size" or "Ticket" in val:
            tooltips.append((addr, "Position size in units. Auto-calculated from risk parameters."))

# ── Apply all tooltips ────────────────────────────────
added = 0
for addr, text in tooltips:
    cell = sh.Range(addr)
    # Delete existing comment if any
    try: cell.Comment.Delete()
    except: pass
    # Add new comment
    cell.AddComment(text)
    cell.Comment.Shape.TextFrame.AutoSize = True
    added += 1

print(f"Added {added} tooltips to Advanced Setup Planner!")

wb.Save()
wb.Close()
excel.Quit()
print("=== VIỆC 9 HOÀN THÀNH ===")
