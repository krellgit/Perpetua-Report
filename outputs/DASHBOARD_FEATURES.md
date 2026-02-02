# Enhanced Excel Dashboard Features

## 📊 What's New in the Enhanced Dashboard

### Fixed Issues
✅ **File Corruption Fixed** - Proper openpyxl formatting, no errors on open
✅ **Embedded Charts** - Charts are IN the Excel file, not external PNG files
✅ **Date Filtering** - Daily data table with Excel's built-in filter dropdowns
✅ **Normalized Dates** - All data aligned to Nov 5, 2025 - Feb 1, 2026

### 4 Interactive Sheets

#### 1. 📊 Executive Dashboard
**What it shows:**
- Performance overview table comparing Perpetua vs Non-Perpetua
- All key metrics: Spend, Sales, Orders, ROAS, ACOS, CPC, Conversion Rate
- Color-coded winners (green = better)
- Embedded ROAS comparison chart

**How to use:**
- Quick glance at overall performance
- Identify which metrics favor Perpetua vs Non-Perpetua
- Visual chart shows ROAS gap immediately

**Key Insights:**
- Non-Perpetua has 2.60 ROAS vs Perpetua's 1.44 (44% better)
- Perpetua drives higher volume ($468K vs $130K sales)
- Non-Perpetua more efficient across all cost metrics

---

#### 2. 📈 Daily Trends
**What it shows:**
- Every single day of data (Nov 5, 2025 - Feb 1, 2026)
- Filterable table with Excel's AutoFilter
- Separate rows for Perpetua and Non-Perpetua each day
- Metrics: Spend, Sales, Orders, Clicks, Impressions, ROAS, ACOS, CPC, CTR, CVR

**How to use:**
1. Click filter dropdown in header row
2. Filter by Date range (select specific weeks/months)
3. Filter by Advertising_Type (view only Perpetua or Non-Perpetua)
4. Analyze trends within your selected period

**Pro Tips:**
- Filter to last 30 days to see recent performance
- Compare weekday vs weekend performance
- Identify seasonal patterns (holidays, promotions)
- Create pivot tables from this data for custom analysis

---

#### 3. 🏆 Top ASINs
**What it shows:**
- Top 50 ASINs by total sales revenue
- Each ASIN color-coded: Blue = Perpetua, Orange = Non-Perpetua
- Metrics: Spend, Sales, Orders, Clicks, ROAS, ACOS

**How to use:**
- Identify star performers (high sales, good ROAS)
- Find optimization opportunities (high spend, low ROAS)
- Compare Perpetua vs Non-Perpetua ASINs side-by-side
- Sort by any column to analyze different angles

**Insights to look for:**
- Are top sellers managed by Perpetua or Non-Perpetua?
- Do high-spending ASINs have acceptable ROAS?
- Which ASINs should get more/less budget?

---

#### 4. 💡 Key Insights
**What it shows:**
- Executive summary of findings
- Primary performance gaps
- Revenue impact analysis
- Optimization opportunities quantified
- Prioritized recommendations

**Strategic Value:**
- **$100K+ opportunity** identified if Perpetua matches non-Perpetua efficiency
- Clear action items with priorities (HIGH/MEDIUM/STRATEGIC)
- Quantified potential impact for each recommendation

---

## 🎨 Dashboard Design Principles Applied

### Best Practices Implemented:

1. **Visual Hierarchy**
   - Sheet names with emojis for quick recognition
   - Color coding: Blue (Perpetua), Orange (Non-Perpetua), Green (Positive), Red (Negative)
   - Clear section headers with larger fonts

2. **Data Accuracy**
   - All dates normalized to common range
   - Proper number formatting ($, %, decimals)
   - No manual entry required

3. **Interactivity**
   - Excel AutoFilter enabled on daily data
   - Can create pivot tables from daily trends
   - Sortable columns

4. **Actionability**
   - Key Insights sheet provides clear next steps
   - Winners identified for each metric
   - Recommendations prioritized by impact

5. **Professional Appearance**
   - Consistent color scheme throughout
   - Proper spacing and alignment
   - Charts embedded (not screenshots)
   - Clean, uncluttered layout

---

## 📅 Date Range Filtering Guide

### Current Date Range:
**November 5, 2025 - February 1, 2026** (89 days)

### How to Filter by Date:

**In Daily Trends sheet:**
1. Click the filter dropdown arrow in "Date" column header
2. Select "Date Filters" → "Between..."
3. Enter your custom date range
4. Click OK

**Example Filters:**
- **Last 30 days:** Jan 3, 2026 - Feb 1, 2026
- **December only:** Dec 1, 2025 - Dec 31, 2025
- **Specific week:** Jan 6, 2026 - Jan 12, 2026

### Multiple Filters:
You can combine filters!
- Filter Date to last 30 days
- AND filter Advertising_Type to "Perpetua"
- Result: See only Perpetua performance for last 30 days

---

## 📊 Chart Insights

### ROAS Comparison Chart (Executive Dashboard)
**Shows:** Side-by-side bar chart comparing Perpetua vs Non-Perpetua ROAS

**What to look for:**
- Height difference = performance gap
- Non-Perpetua bar is 1.8x taller (2.60 vs 1.44 ROAS)
- Visual confirmation of 44% efficiency advantage

**Business Meaning:**
- For every $1 spent:
  - Perpetua generates $1.44 in sales
  - Non-Perpetua generates $2.60 in sales
- Non-Perpetua getting nearly 2x return per dollar

---

## 💡 How to Extract Maximum Value

### For Daily Operations:
1. **Open Daily Trends sheet**
2. Filter to last 7 days
3. Check if ROAS is improving or declining
4. Identify any sudden drops in performance

### For Strategic Planning:
1. **Open Key Insights sheet**
2. Review optimization opportunities
3. Calculate potential ROI of recommendations
4. Prioritize based on HIGH/MEDIUM/STRATEGIC

### For ASIN Optimization:
1. **Open Top ASINs sheet**
2. Sort by ROAS (lowest to highest)
3. Bottom ASINs = optimization targets
4. Review spend levels - pause low performers

### For Executive Reporting:
1. **Open Executive Dashboard**
2. Take screenshot of metrics table
3. Copy ROAS chart to presentation
4. Use for monthly business reviews

---

## 🔄 Updating with New Data

When you get new monthly data:

```bash
cd /mnt/c/Users/Krell/Documents/Imps/gits/Perpetua-Report
# Add new files to data/recent-reports/
python3 scripts/refresh_reports.py
```

This will:
1. Process new data files
2. Recalculate all metrics
3. Generate new dated dashboard
4. Preserve old versions for comparison

**Result:**
- `Perpetua_Dashboard_Enhanced_20260301.xlsx` (March data)
- Compare to `Perpetua_Dashboard_Enhanced_20260202.xlsx` (February data)
- Track month-over-month improvements

---

## 🎯 Key Metrics Explained

### ROAS (Return on Ad Spend)
- **Formula:** Sales ÷ Spend
- **Target:** 2.0+ is good, 3.0+ is excellent
- **Interpretation:** For every $1 spent, how many $ in sales?

### ACOS (Advertising Cost of Sales)
- **Formula:** Spend ÷ Sales
- **Target:** <50% is good, <30% is excellent
- **Interpretation:** What % of sales went to advertising?

### CPC (Cost Per Click)
- **Formula:** Spend ÷ Clicks
- **Target:** Lower is better, depends on category
- **Interpretation:** How much per customer click?

### CVR (Conversion Rate)
- **Formula:** Orders ÷ Clicks
- **Target:** 10%+ is good, 20%+ is excellent
- **Interpretation:** What % of clicks become orders?

---

## 📞 Questions?

See `README.md` for full documentation and usage instructions.

**File Location:**
`C:\Users\Krell\Documents\Imps\gits\Perpetua-Report\outputs\Perpetua_Dashboard_Enhanced_20260202.xlsx`
