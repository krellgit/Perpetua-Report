# Perpetua Campaign Performance Analysis & Reporting

Comprehensive reporting system for analyzing Perpetua advertising campaigns and comparing Perpetua vs non-Perpetua ASIN performance for Nature's Truth Products.

## Overview

This project generates three types of reports:
1. **Campaign Performance Analysis** - Data-driven 4-month performance comparison with visualizations
2. **Executive Summary** - Business narrative documenting the Perpetua automation journey
3. **Technical Documentation** - Implementation reference and reproducibility guide
4. **Excel Dashboard** - Interactive workbook with multiple sheets and metrics

## Quick Start

### 1. Run Full Analysis
```bash
cd /mnt/c/Users/Krell/Documents/Imps/gits/Perpetua-Report
python3 scripts/refresh_reports.py
```

This will:
- Process all campaign data
- Tag ASINs as Perpetua vs Non-Perpetua
- Generate comparison analysis
- Create visualizations
- Build Excel dashboard
- Generate all reports

### 2. View Reports

**Excel Dashboard:** `outputs/Perpetua_Performance_Dashboard_YYYYMMDD.xlsx`
- Executive Summary sheet
- Detailed Comparison sheet
- Top 100 ASINs sheet
- Monthly Trends sheet
- Recommendations sheet

**Markdown Reports:**
- `outputs/Campaign_Performance_Summary.md` - Performance analysis with charts
- `outputs/Executive_Summary_Perpetua_Journey.md` - Full business narrative
- `outputs/Technical_Documentation.md` - Implementation guide

**Visualizations:**
- `outputs/perpetua_vs_nonperpetua_comparison.png`
- `outputs/spend_vs_sales_scatter.png`
- `outputs/roas_comparison.png`
- `outputs/efficiency_metrics_comparison.png`

## Key Findings

### Performance Comparison (4-Month Analysis)

| Metric | Perpetua (227 ASINs) | Non-Perpetua (193 ASINs) |
|--------|----------------------|---------------------------|
| **Total Spend** | $236,157 | $49,596 |
| **Total Sales** | $467,827 | $130,339 |
| **ROAS** | 1.44x | **2.60x** ✓ |
| **ACOS** | 0.10% | **0.06%** ✓ |
| **Avg CPC** | $1.18 | **$0.91** ✓ |

**Insights:**
- Perpetua manages higher-volume, more competitive products (78% of total sales)
- Non-Perpetua shows better efficiency metrics (2.60 ROAS vs 1.44)
- Opportunity to apply non-Perpetua efficiency learnings to Perpetua campaigns

## Project Structure

```
Perpetua-Report/
├── data/
│   ├── recent-reports/          # User-uploaded data files
│   │   ├── SP_Campaign_-_4_Months.csv (19 MB)
│   │   ├── SP_Advertised_Products_-_Max (1).xlsx (14 MB)
│   │   ├── ASIN list - perpetua.xlsx (238 Perpetua ASINs)
│   │   ├── STR_-max_.xlsx (22 MB - Search terms)
│   │   └── SP_Target_Max.xlsx (44 MB)
│   ├── processed/               # Cleaned data
│   │   ├── campaigns_processed.csv
│   │   └── advertised_products_processed.csv
│   └── aggregated/              # Summary tables
│       ├── perpetua_vs_non_perpetua.csv
│       └── asin_level_comparison.json
├── scripts/
│   ├── 1_process_campaign_data.py       # Data processing & tagging
│   ├── 2_asin_level_analysis.py         # ASIN-level comparison
│   ├── 3_generate_performance_report.py # Visualization generation
│   ├── 4_generate_excel_dashboard.py    # Excel workbook creation
│   └── refresh_reports.py               # Automation script (run this!)
├── outputs/
│   ├── Perpetua_Performance_Dashboard_YYYYMMDD.xlsx
│   ├── Campaign_Performance_Report.txt
│   ├── Campaign_Performance_Summary.md
│   ├── Executive_Summary_Perpetua_Journey.md
│   ├── Technical_Documentation.md
│   └── *.png (4 visualizations)
└── README.md (this file)
```

## Data Requirements

### Input Files (Place in `data/recent-reports/`)

1. **ASIN list - perpetua.xlsx** (Required)
   - Sheet 1: "perpetua list" - Perpetua ASINs (ASIN, SKU columns)
   - Sheet 2: "All ASIns" - All ASINs for comparison

2. **SP_Campaign_-_4_Months.csv** (Optional - for campaign-level data)
   - Campaign performance export from Perpetua/Amazon

3. **SP_Advertised_Products_-_Max (1).xlsx** (Recommended)
   - ASIN-level performance data
   - Used for most accurate comparison

4. **STR_-max_.xlsx** (Optional - for search term analysis)
   - Search term report data

5. **SP_Target_Max.xlsx** (Optional - for targeting data)
   - Targeting performance data

### Output Files

All reports are generated in `outputs/` directory:
- Excel dashboard (dated filename for version control)
- Markdown reports (overwritten on each run)
- PNG visualizations (overwritten on each run)
- JSON summaries in `data/aggregated/` (for programmatic access)

## Usage

### One-Command Refresh
```bash
python3 scripts/refresh_reports.py
```

Runs entire pipeline automatically. Safe to re-run with updated data.

### Individual Scripts
```bash
# Step 1: Process and tag campaigns
python3 scripts/1_process_campaign_data.py

# Step 2: ASIN-level analysis
python3 scripts/2_asin_level_analysis.py

# Step 3: Generate visualizations
python3 scripts/3_generate_performance_report.py

# Step 4: Create Excel dashboard
python3 scripts/4_generate_excel_dashboard.py
```

## Dependencies

**Python 3.8+** with:
- pandas (data processing)
- openpyxl (Excel generation)
- matplotlib (visualizations)

**Check installed packages:**
```bash
python3 -c "import pandas, openpyxl, matplotlib; print('All packages available!')"
```

## Perpetua Automation Ecosystem

This report analyzes campaigns managed by a comprehensive 5-project automation system:

1. **Perpetua-Catalog** - Master management hub
2. **Perpetua-Goal-Generator** - Bulk goal creation (12 campaigns/ASIN)
3. **Perpetua-Negative-List** - Automated negative management
4. **Perpetua-SB-SD** - Sponsored Brands/Display automation
5. **Perpetua-Goal-Deleter** - Safe bulk deletion

**Key Achievements:**
- 642 goals under automated management
- 1,140+ hours/year saved through automation
- 95% reduction in manual operations
- 100% API success rate

See `outputs/Executive_Summary_Perpetua_Journey.md` for complete story.

## Recommendations

Based on analysis results:

1. **HIGH PRIORITY:** Investigate non-Perpetua efficiency strategies for application to Perpetua
2. **HIGH PRIORITY:** Focus automation on high-volume while preserving non-Perpetua efficiency
3. **MEDIUM:** Expand negative keyword management to reduce wasted spend
4. **MEDIUM:** Consider migrating top non-Perpetua ASINs to Perpetua
5. **LOW:** Implement automated refresh schedule for ongoing monitoring

## Updating Data

### To Add New Monthly Data:
1. Export latest campaign data from Perpetua/Amazon
2. Save to `data/recent-reports/` with descriptive filename
3. Run `python3 scripts/refresh_reports.py`
4. New dashboard created with updated date

### Comparing Months:
Excel dashboards are dated (YYYYMMDD), allowing month-over-month comparison:
- `Perpetua_Performance_Dashboard_20260201.xlsx` (February)
- `Perpetua_Performance_Dashboard_20260301.xlsx` (March)
- etc.

## Troubleshooting

### "Module not found" errors
```bash
# Check Python packages
python3 -c "import pandas; print('pandas OK')"
python3 -c "import openpyxl; print('openpyxl OK')"
python3 -c "import matplotlib; print('matplotlib OK')"
```

### "File not found" errors
- Ensure all required files are in `data/recent-reports/`
- Check file names match exactly (case-sensitive)

### Script hangs or slow
- Large Excel files (14-44 MB) take time to process
- Allow 2-3 minutes for full pipeline
- Progress indicators shown in terminal

## Support

For questions or issues:
- Review outputs/ for generated reports
- Check data/aggregated/ for JSON summaries
- See scripts/ for Python source code
- GitHub: github.com/krellgit/Perpetua-Report

## License

Internal tool for Nature's Truth Products. Not for redistribution.

---

**Last Updated:** February 2, 2026
