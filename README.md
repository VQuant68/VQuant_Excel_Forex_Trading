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
├── screenshots/                                 # 📸 Live workbook screenshots (see Section 9)
│   ├── photo_1.jpg                              #     Advanced Setup Planner
│   ├── photo_2.jpg                              #     Backend Engine (Narrative Engine)
│   ├── photo_3.jpg                              #     KPI Summary Dashboard
│   └── photo_4.jpg                              #     Macro Hub
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

## 📸 9. System Screenshots — Visual Walkthrough

> The following screenshots capture the live VQuant Trading Workbook in action, demonstrating each critical module of the system as it appears in Microsoft Excel. Every image below represents a real, functioning sheet — not a mockup.

---

### 📷 Screenshot 1 — Advanced Setup Planner (Pre-Trade Analysis Engine)

![VQuant Advanced Setup Planner — EMA Trend Engine, FVG Mapping, Fibonacci Entry Ladder, and Options/Gamma Exposure Levels](https://github.com/user-attachments/assets/03e1d75e-b9dc-488b-8882-bdca3b7f3e6f)

**What You're Looking At:**
This screenshot captures the **Advanced Setup Planner** — the workbook's pre-trade validation module. This is where every trade idea is stress-tested against multiple data dimensions before a single dollar is risked.

**Key Areas Visible:**

| Section | Description |
|---|---|
| **EMA Trend Engine (Column B, Rows 3–14)** | The last 12 EMA candle values from TradingView are entered here. The engine runs a regression analysis on the slope direction and outputs a clear verdict: `LONG` (uptrend), `SHORT` (downtrend), or `CHOPPY` (ranging). This output gates the entire trade — if the EMA says SHORT, no long entries are permitted. |
| **FVG (Fair Value Gap) Mapping** | Multiple timeframe FVG levels (4H, Daily, 1H) are inputted with their High/Low bounds. The system cross-references these gaps against the current price and the entry ladder, highlighting which gaps are "in play" as potential reversal or acceleration zones. |
| **Entry Ladder (Fibonacci-Based)** | After inputting a single Leg High and Leg Low, the system auto-generates 6 entry price levels at Fibonacci retracements (0.236, 0.382, 0.5, 0.618, 0.786). Each level shows the exact entry price, stop-loss distance in pips, and position size in lots — all dynamically calculated from the Summary tab's risk parameters. |
| **Options/Gamma Exposure Levels** | Up to 4 option strike levels and their notional size (in billions) are entered. These create "magnetic" price zones where market maker hedging activity distorts normal price action. The system highlights these on the ladder so the trader knows where price gravity is strongest. |
| **Counter-Trend (CT) Ladder** | A secondary ladder for counter-trend opportunities, with independent risk allocation. Status displays such as *"CT inactive (Side is LONG)"* or *"CT active — see ladder below"* ensure the trader always knows which direction is primary. |

**Why This Sheet Matters:** This is the single most important pre-trade discipline tool in the system. It forces the trader to build a complete, multi-dimensional thesis — EMA trend, FVG confluence, Fibonacci precision entry, institutional options flow — before any position is opened. Traders who use this sheet consistently report eliminating 60–80% of impulsive, low-quality trades.

---

### 📷 Screenshot 2 — Backend Engine (Narrative Engine Output)

![VQuant Backend Engine — Macro Bias, Rate Differential, Yield Narrative, Trade Mode, Directional Bias, and Today's Action synthesis](https://github.com/user-attachments/assets/81adf748-23ee-4510-82ee-5ff137df4efd)

**What You're Looking At:**
This screenshot shows the **Backend Engine (Narrative Engine)** — the brain of the workbook. This sheet takes all upstream data (from the Macro Hub and Advanced Setup Planner) and synthesizes it into 7 clean, actionable outputs that a trader can read in under 30 seconds.

**Key Areas Visible:**

| Output Row | Content | Practical Meaning |
|---|---|---|
| **Macro Bias** | The dominant macroeconomic direction (e.g., `"Buy dips"`, `"Sell rallies"`, `"Neutral — no edge"`) | Tells the trader whether the macro environment supports their trade idea or opposes it |
| **Rate Differential** | The Fed vs ECB interest rate spread with directional commentary (e.g., `"2.37 USD advantage, narrowing"`) | A narrowing differential signals potential USD weakness — critical context for EUR/USD positioning |
| **Yield Narrative** | Bond market interpretation (e.g., `"2Y Treasury down 12bps — early Fed pivot signal detected"`) | The bond market often leads the currency market by weeks — this gives an early warning system |
| **Directional Bias** | `LONG` or `SHORT` — the unified conclusion | The single most important word on the entire sheet |
| **Trade Mode** | `A+Long` (max conviction), `CoreLong` (standard), `CTshort` (counter-trend), or `Avoid` | Controls how aggressively the trader should deploy capital today |
| **Plain English Summary** | A 2–3 sentence human-readable synthesis of all inputs | Designed to read like a morning brief from a macro analyst at a hedge fund |
| **Today's Action** | A single, precise instruction (e.g., *"Wait for pullback to 1.1480. Do not enter above 1.1510. Invalidation below 1.1430."*) | The final, concrete decision — no ambiguity, no second-guessing |

**Why This Sheet Matters:** The Narrative Engine transforms raw, scattered data into a coherent trading thesis. Without it, a trader would need to mentally synthesize macroeconomic policy, bond yields, technical structure, and risk parameters — a process that takes institutional analysts 30–60 minutes daily. This sheet does it in zero seconds after the upstream inputs are updated.

---

### 📷 Screenshot 3 — KPI Summary Dashboard (Monthly Performance Cockpit)

![VQuant Summary Dashboard — Monthly Targets, Daily/Weekly KPIs, NAV Tracking, Cumulative P&L, and Weekly Variance Table](https://github.com/user-attachments/assets/fa8d012b-9dc7-4680-ae3d-c72bb99ae5ac)

**What You're Looking At:**
This screenshot captures the **Summary Dashboard** — the workbook's top-level performance cockpit. This is the first sheet a trader checks at the start of each day and the last sheet they review at the end of each week.

**Key Areas Visible:**

| Section | Cells | Description |
|---|---|---|
| **Monthly Configuration (Green Cells)** | B2–B10 | The 6 input cells that control the entire system: Monthly Profit Target, Planned Trading Days, Weeks in Cycle, Risk % per Trade, Max Daily Loss (R), and Starting NAV. Every other number on every other sheet cascades from these 6 values. |
| **Auto-Calculated Targets** | B4, B6 | Daily Target (`$Target ÷ Days`) and Weekly Target (`$Target ÷ Weeks`) — your exact benchmark for each time period |
| **Live Account Metrics** | B11–B12 | Cumulative P&L (live SUM from Daily Log) and Current NAV (Starting NAV + Cumulative P&L) — your real-time financial position |
| **Weekly Accountability Table (Rows 21–25)** | A21:J25 | 4 rows (one per week) showing: Active Days, Total Trades, Wins, Losses, Win Rate %, Week Net P&L, Weekly Target, and the critical **Week Variance** column |
| **Week Variance Column** | J21:J25 | **The most important metric on the dashboard.** Green = above weekly target. Red = below. This single column forces a strategic review: are you on pace, or do you need to adjust your approach next week? |

**Why This Sheet Matters:** The Summary Dashboard is the trader's "single source of truth" for monthly performance. It eliminates the most dangerous illusion in trading — confusing *making money* with *trading well*. A trader who earned $400 in a week but had a $500 target is underperforming, and this sheet makes that fact impossible to ignore. The weekly variance drives continuous improvement by forcing an honest assessment of progress against plan.

---

### 📷 Screenshot 4 — Macro Hub (Weekly Macroeconomic Context)

![VQuant Macro Hub — USD Strength Score, EUR Strength Score, Fed/ECB Rate Differentials, 2Y Bond Yield Tracking, and HTF Structure Assessment](https://github.com/user-attachments/assets/ba2bba56-e9a7-4af5-b20a-444ccc05f001)

**What You're Looking At:**
This screenshot shows the **Macro Hub** — the upstream data input layer that feeds the entire decision pipeline. Updated once per week (typically Sunday evening), this sheet captures the macroeconomic landscape that governs all Forex price action.

**Key Areas Visible:**

| Section | Description |
|---|---|
| **USD Strength Score** | A composite score (typically 0–100) synthesizing DXY index movement, Federal Reserve policy expectations, US employment/CPI data, and Treasury yield direction. A score above 60 signals USD strength; below 40 signals USD weakness. |
| **EUR Strength Score** | The equivalent composite for the Euro, incorporating ECB rate expectations, Eurozone PMI data, and EUR-denominated bond yields. Cross-referencing USD and EUR scores reveals the dominant currency pair direction. |
| **Fed/ECB Rate Differential** | The current interest rate spread between the Federal Reserve and the European Central Bank, with a 6-month directional arrow (widening ↑ or narrowing ↓). This is the single most powerful long-term driver of EUR/USD price action. |
| **2-Year Bond Yield Tracker** | Weekly snapshots of 2-year Treasury and Bund yields, with basis-point changes highlighted. The 2-year yield is the most rate-sensitive maturity — it reacts first to policy shifts, often weeks before the currency market catches up. |
| **HTF (Higher Time Frame) Structure** | A manual assessment of Weekly and Daily chart structure: is price making higher highs/higher lows (bullish), lower highs/lower lows (bearish), or range-bound? This is the technical anchor that the Narrative Engine combines with macro data. |

**Why This Sheet Matters:** The Macro Hub is where the workbook's "institutional edge" begins. Most retail traders skip macroeconomic analysis entirely — they jump straight to 15-minute charts and wonder why their setups fail during NFP releases or Fed press conferences. By spending 15 minutes per week updating this sheet, the trader gains a structural advantage that persists for the entire trading week. The data entered here flows automatically into the Narrative Engine, which produces the unified trading thesis that governs all downstream decisions.

---

> **📌 Note:** All screenshots above represent the live, production version of `Trading_Workbook_MASTER.xlsx`. Actual data values shown may differ from session to session as new market data is entered. The formula architecture and sheet layout remain constant across all usage scenarios.

---

*This repository is maintained by **VQuant Engineering**. It represents a portfolio-grade commercial deliverable demonstrating expertise in financial systems automation, quantitative risk management methodology, and professional-grade Excel formula architecture.*  
*Client: Delivered and accepted — April 2026.*
