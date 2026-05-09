# 📊 VQuant — Forex Trading Master Workbook

> **An Institutional-Grade, Fully Automated Forex Analysis & Trade Management System built in Microsoft Excel — engineered for professional retail traders who demand data-driven decisions, not gut feelings.**

---

## 🧭 1. Why This Workbook Exists — The Problem We Solved

Every retail Forex trader faces the same three invisible killers:

### ❌ Problem 1: Emotional Decision-Making
You open your charts, you *feel* like the market is going up. You enter. You lose. The trade wasn't based on data — it was based on bias. Most retail traders have no system that forces them to validate a trade idea *before* risking money. They rely on gut instinct, which is statistically unreliable over thousands of trades.

### ❌ Problem 2: Inconsistent Risk Management
One day you risk $50. The next day, after a big win, you risk $200 on the same setup. This inconsistency means even a profitable strategy can blow an account. Without hard-coded, formula-driven position sizing, emotion always creeps back in.

### ❌ Problem 3: No Post-Trade Accountability
After a losing week, most traders have no idea *why* they lost. Was it bad setups? Bad execution? Bad timing? Without structured journaling linked to actual performance data, there is no learning — only repeating the same mistakes.

---

**VQuant Engineering built this workbook to systematically eliminate all three problems.** It replaces gut feeling with a mathematical pipeline that takes a trader from macro context all the way to a single, decisive, risk-calculated trade instruction — automatically.

---

## 🏗️ 2. The 6-Sheet System Architecture

The workbook is not just a spreadsheet. It is a **6-module decision engine** where every sheet has a specific, non-overlapping role. Data flows automatically downstream through `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, and linked formula chains. The trader never needs to copy-paste between sheets.

```text
┌──────────────────────────────────────────────────────┐
│  SHEET 1: Macro Hub                                  │
│  Weekly macroeconomic context: Fed rates, bond       │
│  yields, USD/EUR strength scores.                    │
└────────────────────┬─────────────────────────────────┘
                     │ feeds
┌────────────────────▼─────────────────────────────────┐
│  SHEET 2: Advanced Setup Planner                     │
│  Pre-trade technical analysis: EMA, FVG, Structure,  │
│  Options/Gamma levels, Entry Ladder.                 │
└────────────────────┬─────────────────────────────────┘
                     │ feeds
┌────────────────────▼─────────────────────────────────┐
│  SHEET 3: Backend Engine (Narrative Engine)          │
│  Synthesizes all inputs → outputs 7 structured       │
│  decisions: Bias, Mode, Action, Narrative.           │
└────────────────────┬─────────────────────────────────┘
                     │ confirmed decision flows into
┌────────────────────▼─────────────────────────────────┐
│  SHEET 4: Raw Daily Data                             │
│  Paste MT4/MT5 export here. Trade Date and           │
│  Cumulative P&L auto-populate.                       │
└────────────────────┬─────────────────────────────────┘
                     │ feeds
┌────────────────────▼─────────────────────────────────┐
│  SHEET 5: Daily Log                                  │
│  End-of-day reconciliation: Win Rate, R-Multiple,    │
│  Net P&L, Week number assignment.                    │
└────────────────────┬─────────────────────────────────┘
                     │ aggregates into
┌────────────────────▼─────────────────────────────────┐
│  SHEET 6: Summary                                    │
│  Monthly KPI Dashboard: NAV, Weekly Variance,        │
│  Progress to Target, Required R/week.                │
└──────────────────────────────────────────────────────┘
```

**Why this matters to you as a client:** Every number you see on the Summary dashboard is 100% derived from your actual trade data. You cannot lie to yourself. The system forces an objective reality check every single week.

---

## 💡 3. Feature-by-Feature Breakdown (Client Value Focus)

---

### 🧠 Feature 1: The Narrative Engine — Your Personal Macro Analyst

**The Problem It Solves:** Most retail traders ignore macroeconomics because it's too complex and time-consuming to synthesize manually. They end up trading against the dominant macro theme and wonder why their technically-perfect setups keep failing.

**What It Does:** The Narrative Engine ingests 5 separate data streams — USD strength score, EUR strength score, Fed/ECB interest rate differential, 2-year bond yield movements, and HTF price structure — and synthesizes them into 7 clean, structured outputs:

| Output | What It Tells You | Real Example |
|---|---|---|
| **Macro Bias** | Which direction does the macroeconomic environment favor? | `"Buy dips"` — macro supports USD longs |
| **Rate Differential** | Is the interest rate gap between Fed and ECB widening or narrowing? | `"2.37 USD advantage, narrowing over 6 months"` |
| **Yield Narrative** | What is the bond market signaling about future policy? | `"2Y Treasury down 12bps — early Fed pivot signal detected"` |
| **Directional Bias** | What is the overall directional lean combining macro AND structure? | `LONG` or `SHORT` |
| **Trade Mode** | How aggressively should you be trading today? | `A+Long` (max conviction) / `CTshort` / `CoreLong` / `Avoid` |
| **Plain English Summary** | A 2–3 sentence synthesis a human trader can act on immediately | *"Macro supports EUR longs. USD advantage narrowing on Fed cut expectations. HTF structure bullish — target 0.382–0.618 Fibonacci retracement."* |
| **Today's Action** | A single, precise instruction: what to do right now | *"Wait for pullback to 1.1480. Do not enter above 1.1510. Invalidation below 1.1430."* |

**Why this matters:** You no longer need to spend 45 minutes reading ForexFactory and CME FedWatch data. You update the Macro Hub sheet weekly (15 minutes), and the engine instantly tells you the unified conclusion. The Narrative Engine is designed to sound like — and reason like — a seasoned macro analyst, not a generic AI chatbot.

---

### 📐 Feature 2: Advanced Setup Planner — Pre-Trade Validation Checklist

**The Problem It Solves:** Most traders enter trades based on a single signal — one candle pattern, one EMA cross. This ignores the broader confluence of factors that separate high-probability setups from noise. The Advanced Setup Planner forces a full, multi-dimensional validation before any trade is considered.

**What It Does:**

#### 2a. EMA Trend Engine (12-row input)
- You input the last 12 EMA candle values from TradingView into Column B.
- The engine performs a regression analysis on the EMA slope direction.
- Output: `LONG` (EMA fanning upward) / `SHORT` (EMA fanning downward) / `CHOPPY` (flat/overlapping).
- **Why it matters:** You never again enter a long trade into a downtrending EMA. The system physically blocks that decision.

#### 2b. FVG (Fair Value Gap) Mapping
- Input 4H, Daily, and 1H Fair Value Gap levels (high and low of each gap).
- The engine cross-references these against the current price and the Entry Ladder.
- **Why it matters:** FVGs represent institutional imbalances — price magnets that Smart Money targets. Knowing where these are in advance tells you exactly where price is likely to pause, reverse, or accelerate.

#### 2c. Entry Ladder (Fibonacci-Based Position Builder)
- Input a single Leg High and Leg Low.
- The system auto-generates a full Entry Ladder: 6 entry price levels (at 0.236, 0.382, 0.5, 0.618, 0.786 Fibonacci), with stop-loss and position size in lots calculated at each level.
- **Why it matters:** Instead of putting your entire position on one candle, you scale in at mathematically optimal levels — reducing average entry price on longs and maximizing the Risk/Reward of the overall position.

#### 2d. Options/Gamma Exposure Levels
- Input up to 4 Option Strike levels and their size in billions (from CME or your options flow source).
- The system highlights these on the ladder as key magnetic price levels.
- **Why it matters:** Market makers are obligated to hedge their gamma exposure. This creates predictable price gravity around certain strikes that technical analysis alone cannot see.

#### 2e. Counter-Trend (CT) Ladder
- When the primary bias is `LONG`, the system auto-generates a separate CT ladder for short setups that meet the threshold.
- Displays a live status message: *"CT inactive (Side is LONG)."* or *"CT active — see ladder below."*
- **Why it matters:** Top-tier traders don't ignore counter-trend opportunities — they manage them with separate, smaller risk allocations. This feature enables that discipline without manual recalculation.

---

### 📋 Feature 3: Automated Trade Journaling — Zero Manual Calculation

**The Problem It Solves:** Most traders hate journaling because it's tedious and manual. They start doing it for two weeks, then quit. Without consistent journaling, there is no data to improve from.

**What It Does:** The Raw Daily Data and Daily Log sheets create a completely automated journaling pipeline:

**Step 1 — Raw Data (30 seconds of work):**
- After closing your trading session, export your trade history from MT4/MT5 as a standard file.
- Open the `Raw Daily Data` tab. Click Cell A2. Paste.
- `Column I (Trade Date)`: Auto-extracts and normalizes the timestamp from the broker's format.
- `Column J (Cumulative P&L)`: Instantly shows your running P&L total from the very first trade of the month.

**Step 2 — Daily Log (3 cells to fill):**
- Enter the date in A2.
- Enter the week number in B2 (Week 1, 2, 3, or 4).
- Type `Y` in D2 (to mark that you traded today).
- The system auto-calculates:
  - Net P&L for the day in dollars.
  - Win Rate % for today's session.
  - R-Multiple: Did you earn what you risked, or less?

**Step 3 — Nothing:**
- The Summary dashboard updates itself automatically from the Daily Log.

**Why it matters:** Journaling is now a 3-minute end-of-day ritual, not a 45-minute chore. Traders who journal consistently outperform those who don't — this system eliminates every friction point that makes people quit journaling.

---

### 📈 Feature 4: KPI Summary Dashboard — Monthly Performance Cockpit

**The Problem It Solves:** Most traders don't know if they're actually profitable on a risk-adjusted basis. They confuse a "good month" (made money) with a "good trading month" (made money *efficiently*, with controlled drawdown and improving Win Rate). The Summary dashboard forces this distinction.

**How To Use It (One-Time Monthly Setup — 5 Minutes):**
Open the `Summary` tab at the start of each month and fill in only the green-highlighted cells:

| Cell | Input | Example | Why It Matters |
|---|---|---|---|
| **B2** | Monthly Profit Target | `$2,000` | Your north star for the month |
| **B3** | Planned Trading Days | `20 days` | Sets your daily target automatically |
| **B5** | Weeks in Cycle | `4 weeks` | Sets your weekly target automatically |
| **B8** | Risk % per Trade | `5%` (= 0.05) | Controls your position size engine |
| **B9** | Max Daily Loss (R) | `2R` | Hard stop: lose 2R → stop trading today |
| **B10** | Starting NAV | `$10,000` | Your baseline capital for the month |

**What Auto-Updates Every Single Day:**

- **`B4 — Daily Target`:** `$2,000 / 20 = $100/day` — the exact dollar amount you need to earn each trading day to hit your monthly goal.
- **`B6 — Weekly Target`:** `$2,000 / 4 = $500/week` — your weekly KPI checkpoint.
- **`B11 — Cumulative P&L`:** A live `SUM` of every row in Daily Log's Net P&L column. This is your real-time profit for the month.
- **`B12 — Current NAV`:** `Starting NAV + Cumulative P&L`. Your account size at this exact moment.

**The Weekly Table (Rows 21–25) — Your Accountability Mirror:**
Each row represents one week of the month. Every cell is fully automated via `SUMIFS` and `COUNTIFS`:

| Column | Auto-Calculated Value |
|---|---|
| Active Days | How many days you actually traded this week |
| Total Trades | Total number of positions opened |
| Wins / Losses | Breakdown by outcome |
| Win Rate % | `Wins ÷ Total Trades` — your decision quality score |
| Week Net P&L | Exact dollar profit/loss for this week |
| Weekly Target | Your pre-set target ($500) for comparison |
| **Week Variance** | **Green = above target. Red = below target.** |

**Why the Variance column is the most important number on the sheet:** It separates *how much you made* from *whether you're on track*. A trader can make $400 in a week and still be underperforming if their target was $500. The variance forces a weekly strategic review: do you need to trade more days next week, or improve your Win Rate?

---

## 🛡️ 5. Risk Management — Hard Rules Encoded in Formulas

The system doesn't rely on willpower. Risk rules are physically encoded at the formula level:

### Rule 1: Max Daily Loss Hard Stop
When the Daily Log detects that a trader has lost `Max Daily Loss` (e.g., 2R in one day), the sheet visually flags the day. This is the system's way of enforcing the most important rule in trading: **stop trading when your edge is compromised for the day.**

### Rule 2: Dynamic Position Sizing (Auto-Compounding)
Position sizes are calculated as a **percentage of the Current NAV**, not a fixed dollar amount. This means:
- After a winning month ($10,000 → $11,000), position sizes grow proportionally — compounding gains.
- After a losing streak ($10,000 → $9,200), position sizes shrink automatically — protecting remaining capital.
- A trader cannot "revenge trade" their way into a larger hole. The system reduces risk exposure precisely when emotions are highest.

### Rule 3: No Hard-Coded Numbers Anywhere
Every risk parameter cascades from a single source of truth — the 6 green cells in the Summary tab. Change one number there, and the position size engine, daily target, weekly target, and R-multiple calculations all update instantaneously across every sheet.

---

## 📂 6. Repository Contents

```text
.
├── Trading_Workbook_MASTER.xlsx                 # 📊 The complete 6-sheet trading system
├── Quant_Excel_Report.html                      # 📄 HTML delivery report with full feature documentation
├── USER_MANUAL_SUMMARY_REPORT.txt               # 📚 Admin-level formula anatomy documentation
├── CLIENT_TEST_SCENARIO_EN.txt                  # 🧪 Step-by-step guided walkthrough for first-time users
├── VQuant — Forex Trading Workbook · Final Delivery Report.pdf  # 📑 Professional client handover document
└── README.md                                    # 📖 Full system documentation (you are here)
```

---

## 🚀 7. Quick Start Guide (First Time Use — 15 Minutes)

```
STEP 1: Open Trading_Workbook_MASTER.xlsx

STEP 2: Go to "Macro Hub" tab
        → Update bond yields, rate differentials, and USD/EUR scores (weekly, ~15 min)
        → The Narrative Engine auto-generates your weekly trade bias.

STEP 3: Go to "Advanced Setup Planner"
        → Input the last 12 EMA values from your TradingView chart
        → Input your Leg High / Leg Low for the current setup
        → The Entry Ladder appears instantly with pre-calculated sizes and stops.

STEP 4: Execute your trade on MT4/MT5

STEP 5: After closing → Export trade history from MT4/MT5
        → Paste into "Raw Daily Data" tab → Done.
        → Open "Daily Log" → Mark 3 cells → Done.
        → Check "Summary" → Your NAV, Win Rate, and Weekly Variance updated automatically.
```

> **No manual math. No copy-pasting between sheets. No forgetting to update a formula.  
> The entire pipeline from macro context to final P&L report is 100% automated.**

---

## 🏆 8. Technical Delivery Specifications

| Specification | Detail |
|---|---|
| **Platform** | Microsoft Excel `.xlsx` — compatible with Excel 2019, Excel 365, and Excel for Mac |
| **Data Source (Macro)** | ForexFactory economic calendar + CME FedWatch tool (updated manually weekly) |
| **Data Source (Trades)** | MT4 / MT5 standard trade history export (CSV paste format) |
| **Data Source (Technical)** | TradingView charts (EMA and FVG values entered manually before each session) |
| **Formula Complexity** | Advanced `SUMIFS`, `COUNTIFS`, `IFERROR`, `INDEX/MATCH`, dynamic string concatenation |
| **Sheet Protection** | All formula cells are sheet-protected; only designated input cells are unlocked |
| **Performance** | All formulas resolve in real-time with no perceptible lag across all 6 sheets |
| **Logging Architecture** | 100-row capacity per month in Daily Log (expandable by copying formula rows down) |

---

*This repository is maintained by **VQuant Engineering**. It represents a portfolio-grade commercial deliverable demonstrating expertise in financial systems automation, quantitative risk management methodology, and professional-grade Excel formula architecture.*  
*Client: Delivered and accepted — April 2026.*
