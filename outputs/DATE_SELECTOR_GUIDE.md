# 📅 Date Selector Feature Guide

## ✅ Your Dashboard Now Has TWO Ways to Filter by Date

### METHOD 1: Date Dropdown Selectors (Dashboard Sheet)

**Location:** `📊 Dashboard` sheet, cells **D6** and **G6**

**How to Use:**
1. Open the dashboard
2. Click on cell **D6** (Start Date)
3. You'll see a small dropdown arrow appear on the right side of the cell
4. Click the dropdown arrow
5. Scroll through the list of dates and select your start date
6. Repeat for cell **G6** (End Date)

**What This Does:**
- Shows you which date range you're analyzing
- Visual reference for your selected period
- Can manually type dates too (format: YYYY-MM-DD)

**Note:** The metrics on the Dashboard sheet show ALL data. For filtered metrics, use Method 2.

---

### METHOD 2: AutoFilter (Daily Data Sheet) - **RECOMMENDED**

**Location:** `📅 Daily Data (Filter Here)` sheet

**How to Use:**
1. Go to "📅 Daily Data" sheet
2. Look at the header row - you'll see small dropdown arrows ▼
3. Click the dropdown arrow in the **"Date"** column
4. Select one of these options:

**Option A: Filter to Specific Dates**
- Uncheck "(Select All)"
- Check only the dates you want
- Click OK

**Option B: Date Range Filter**
- Hover over "Date Filters"
- Select "Between..."
- Enter Start Date and End Date
- Click OK

**Option C: Relative Filters**
- "Last Week" - Shows last 7 days
- "Last Month" - Shows last 30-31 days
- "This Month" - Current month only
- Custom ranges as needed

**What This Does:**
- Instantly filters all data to your date range
- Only shows rows matching your criteria
- Can combine with other filters (e.g., "Perpetua" only)
- Excel automatically recalculates totals if you use SUM formulas

---

## 📊 Understanding the Data

### Normalized Date Range
**Full Dataset:** November 5, 2025 - February 1, 2026 (89 days)

This date range covers data from ALL your source files:
- 4-month campaign report
- Advertised products report
- All dates normalized to this common range

### Why Two Methods?

**Method 1 (Dropdowns):**
- Good for documentation
- Shows selected range clearly
- Easy to see at a glance

**Method 2 (AutoFilter):**
- Actually filters the data
- Instant visual feedback
- Can see filtered totals
- More powerful for analysis

---

## 💡 Example Workflows

### Analyze Last 30 Days Performance
1. Go to "Daily Data" sheet
2. Click Date filter dropdown
3. Select "Date Filters" > "Between..."
4. Start: 2026-01-03, End: 2026-02-01
5. Review filtered data

### Compare Weekdays vs Weekends
1. Filter to weekdays (Mon-Fri)
2. Note the ROAS values
3. Clear filter
4. Filter to weekends (Sat-Sun)
5. Compare performance

### Perpetua Performance Last Week
1. Filter Date to last 7 days
2. Filter Advertising_Type to "Perpetua"
3. Review metrics

---

## 📈 What the Correct Data Shows

### Aggregate Metrics (Entire Period):

**Perpetua (SaaS) - 227 ASINs:**
- Spend: $236,157
- Sales: $467,827
- **ROAS: 1.98x**
- **ACOS: 50.5%**
- CPC: $1.18
- CTR: 0.68%
- CVR: 19.84%

**Non-Perpetua (Manual) - 193 ASINs:**
- Spend: $49,596
- Sales: $130,339
- **ROAS: 2.63x**
- **ACOS: 38.1%**
- CPC: $0.91
- CTR: 0.73%
- CVR: 20.85%

### Winner: Non-Perpetua (33% Better ROAS)

**What This Means:**
- For every $1 spent, Non-Perpetua returns $2.63 vs Perpetua's $1.98
- Non-Perpetua is more cost-efficient across ALL metrics
- Perpetua drives higher volume but lower efficiency

**Potential Opportunity:**
- If Perpetua matched Non-Perpetua ROAS: +$154,000 additional revenue
- This is your optimization target!

---

## 🔍 Finding the Date Selectors

**In Excel:**
1. Open: `Perpetua_Dashboard_with_DateSelector_20260202_0919.xlsx`
2. Look at cells **D6** and **G6** on the Dashboard sheet
3. Click on these cells to see the dropdown arrow appear
4. The dropdown contains all 89 dates in your dataset

**Visual Indicator:**
- Cells D6 and G6 have **yellow background**
- This indicates they are input/selector cells
- Bold text shows they are interactive

---

## ✅ All Metrics Included

Your dashboard includes ALL valuable testing metrics:

**Efficiency Metrics:**
1. ✅ ROAS (Return on Ad Spend)
2. ✅ ACOS (Advertising Cost of Sales)
3. ✅ CPC (Cost Per Click)
4. ✅ CTR (Click-Through Rate)
5. ✅ CVR (Conversion Rate)
6. ✅ CPA (Cost Per Acquisition)
7. ✅ CPM (Cost Per Mille)
8. ✅ AOV (Average Order Value)

**Volume Metrics:**
9. ✅ Total Spend
10. ✅ Total Sales
11. ✅ Total Orders
12. ✅ Total Clicks
13. ✅ Total Impressions
14. ✅ Total Units

**Productivity Metrics:**
15. ✅ Spend per ASIN
16. ✅ Sales per ASIN
17. ✅ Orders per ASIN

**Every metric is calculated correctly using aggregate portfolio-level calculations!**
