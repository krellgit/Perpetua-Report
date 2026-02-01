#!/usr/bin/env python3
"""
Process 4-month campaign data and compare Perpetua vs non-Perpetua performance
"""

import pandas as pd
import re
import json
from pathlib import Path
from datetime import datetime

# Paths
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / 'data' / 'recent-reports'
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
AGG_DIR = BASE_DIR / 'data' / 'aggregated'

# Create output directories
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
AGG_DIR.mkdir(parents=True, exist_ok=True)

print("=" * 80)
print("PERPETUA VS NON-PERPETUA CAMPAIGN ANALYSIS")
print("=" * 80)
print()

# Step 1: Load Perpetua ASINs (238 ASINs)
print("[1/6] Loading Perpetua ASIN list...")
perpetua_df = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='perpetua list')
perpetua_asins = set(perpetua_df['ASIN'].dropna().str.strip().tolist())
print(f"  ✓ Loaded {len(perpetua_asins)} Perpetua ASINs")

# Step 2: Load all ASINs (455 total)
print("[2/6] Loading all ASINs...")
all_asins_df = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='All ASIns')
all_asins = set(all_asins_df['ASIN (Informational only)'].dropna().str.strip().tolist())
print(f"  ✓ Loaded {len(all_asins)} total ASINs")

# Calculate non-Perpetua ASINs
non_perpetua_asins = all_asins - perpetua_asins
print(f"  ✓ Calculated {len(non_perpetua_asins)} non-Perpetua ASINs")
print()

# Step 3: Load campaign data
print("[3/6] Loading 4-month campaign data...")
campaigns = pd.read_csv(DATA_DIR / 'SP_Campaign_-_4_Months.csv')
print(f"  ✓ Loaded {len(campaigns):,} campaign records")
print(f"  ✓ Unique campaigns: {campaigns['Campaign Name'].nunique():,}")
print(f"  ✓ Date range: {campaigns['Date'].min()} to {campaigns['Date'].max()}")
print()

# Step 4: Extract ASIN from campaign names
print("[4/6] Extracting ASINs from campaign names...")

def extract_asin(campaign_name):
    """Extract ASIN (B0XXXXXXXXX format) from campaign name"""
    if pd.isna(campaign_name):
        return None
    # Look for B0 followed by 8 alphanumeric characters
    match = re.search(r'B[A-Z0-9]{9}', str(campaign_name))
    return match.group(0) if match else None

campaigns['ASIN'] = campaigns['Campaign Name'].apply(extract_asin)
campaigns_with_asin = campaigns[campaigns['ASIN'].notna()]
print(f"  ✓ Extracted ASINs from {len(campaigns_with_asin):,} campaigns ({len(campaigns_with_asin)/len(campaigns)*100:.1f}%)")
print(f"  ✓ Unique ASINs found: {campaigns_with_asin['ASIN'].nunique()}")
print()

# Step 5: Tag campaigns as Perpetua vs Non-Perpetua
print("[5/6] Tagging campaigns as Perpetua vs Non-Perpetua...")

def classify_campaign(asin):
    """Classify campaign based on ASIN"""
    if pd.isna(asin):
        return 'Unknown'
    if asin in perpetua_asins:
        return 'Perpetua'
    elif asin in non_perpetua_asins:
        return 'Non-Perpetua'
    else:
        return 'Unknown'

campaigns['Advertising_Type'] = campaigns['ASIN'].apply(classify_campaign)

# Distribution
type_counts = campaigns['Advertising_Type'].value_counts()
print(f"  Campaign distribution:")
for ad_type, count in type_counts.items():
    pct = count / len(campaigns) * 100
    print(f"    {ad_type:15s}: {count:7,} ({pct:5.1f}%)")
print()

# Filter to known campaigns only
known_campaigns = campaigns[campaigns['Advertising_Type'].isin(['Perpetua', 'Non-Perpetua'])].copy()
print(f"  ✓ Analyzing {len(known_campaigns):,} campaigns with known ASINs")
print()

# Clean numeric columns
print("[6/6] Processing metrics...")
numeric_cols = ['Spend', 'Cost Per Click (CPC)', 'Impressions', 'Clicks',
                '7 Day Total Orders (#)', '7 Day Total Sales ']

for col in numeric_cols:
    if col in known_campaigns.columns:
        # Remove $ and % symbols, convert to numeric
        if known_campaigns[col].dtype == 'object':
            known_campaigns[col] = known_campaigns[col].astype(str).str.replace('$', '').str.replace('%', '').str.replace(',', '')
        known_campaigns[col] = pd.to_numeric(known_campaigns[col], errors='coerce').fillna(0)

# Calculate ACOS and ROAS if not present or if they need cleaning
if 'Total Advertising Cost of Sales (ACOS) ' in known_campaigns.columns:
    known_campaigns['ACOS'] = pd.to_numeric(
        known_campaigns['Total Advertising Cost of Sales (ACOS) '].astype(str).str.replace('%', ''),
        errors='coerce'
    ) / 100
else:
    # Calculate ACOS = Spend / Sales
    known_campaigns['ACOS'] = known_campaigns.apply(
        lambda x: x['Spend'] / x['7 Day Total Sales '] if x['7 Day Total Sales '] > 0 else 0,
        axis=1
    )

if 'Total Return on Advertising Spend (ROAS)' in known_campaigns.columns:
    known_campaigns['ROAS'] = pd.to_numeric(
        known_campaigns['Total Return on Advertising Spend (ROAS)'],
        errors='coerce'
    ).fillna(0)
else:
    # Calculate ROAS = Sales / Spend
    known_campaigns['ROAS'] = known_campaigns.apply(
        lambda x: x['7 Day Total Sales '] / x['Spend'] if x['Spend'] > 0 else 0,
        axis=1
    )

# Calculate CTR if needed
if 'Click-Thru Rate (CTR)' in known_campaigns.columns:
    known_campaigns['CTR'] = pd.to_numeric(
        known_campaigns['Click-Thru Rate (CTR)'].astype(str).str.replace('%', ''),
        errors='coerce'
    ) / 100
else:
    known_campaigns['CTR'] = known_campaigns.apply(
        lambda x: x['Clicks'] / x['Impressions'] if x['Impressions'] > 0 else 0,
        axis=1
    )

print("  ✓ Metrics cleaned and calculated")
print()

# Save processed data
processed_file = PROCESSED_DIR / 'campaigns_processed.csv'
known_campaigns.to_csv(processed_file, index=False)
print(f"✓ Saved processed data to: {processed_file}")
print()

# Aggregate by Advertising Type
print("=" * 80)
print("PERPETUA VS NON-PERPETUA COMPARISON")
print("=" * 80)
print()

comparison = known_campaigns.groupby('Advertising_Type').agg({
    'Spend': 'sum',
    '7 Day Total Sales ': 'sum',
    '7 Day Total Orders (#)': 'sum',
    'Impressions': 'sum',
    'Clicks': 'sum',
    'Campaign Name': 'count'  # Number of campaign records
}).round(2)

# Rename columns
comparison.columns = ['Total_Spend', 'Total_Sales', 'Total_Orders',
                      'Total_Impressions', 'Total_Clicks', 'Campaign_Records']

# Calculate averages
avg_metrics = known_campaigns.groupby('Advertising_Type').agg({
    'ACOS': 'mean',
    'ROAS': 'mean',
    'Cost Per Click (CPC)': 'mean',
    'CTR': 'mean'
}).round(4)

# Combine
comparison = comparison.join(avg_metrics)

# Calculate derived metrics
comparison['Avg_CPC'] = comparison['Total_Spend'] / comparison['Total_Clicks']
comparison['Conversion_Rate'] = comparison['Total_Orders'] / comparison['Total_Clicks']

print(comparison.to_string())
print()

# Calculate deltas
if 'Perpetua' in comparison.index and 'Non-Perpetua' in comparison.index:
    perpetua = comparison.loc['Perpetua']
    non_perpetua = comparison.loc['Non-Perpetua']

    print("=" * 80)
    print("PERFORMANCE DELTA (Perpetua vs Non-Perpetua)")
    print("=" * 80)
    print()

    deltas = {
        'Total Spend': perpetua['Total_Spend'] - non_perpetua['Total_Spend'],
        'Total Sales': perpetua['Total_Sales'] - non_perpetua['Total_Sales'],
        'Total Orders': perpetua['Total_Orders'] - non_perpetua['Total_Orders'],
        'Avg ACOS': perpetua['ACOS'] - non_perpetua['ACOS'],
        'Avg ROAS': perpetua['ROAS'] - non_perpetua['ROAS'],
        'Avg CPC': perpetua['Avg_CPC'] - non_perpetua['Avg_CPC'],
        'Conversion Rate': perpetua['Conversion_Rate'] - non_perpetua['Conversion_Rate']
    }

    # Calculate percentage improvements
    improvements = {
        'ACOS Improvement': ((non_perpetua['ACOS'] - perpetua['ACOS']) / non_perpetua['ACOS'] * 100) if non_perpetua['ACOS'] > 0 else 0,
        'ROAS Improvement': ((perpetua['ROAS'] - non_perpetua['ROAS']) / non_perpetua['ROAS'] * 100) if non_perpetua['ROAS'] > 0 else 0,
        'CPC Improvement': ((non_perpetua['Avg_CPC'] - perpetua['Avg_CPC']) / non_perpetua['Avg_CPC'] * 100) if non_perpetua['Avg_CPC'] > 0 else 0,
        'Conversion Improvement': ((perpetua['Conversion_Rate'] - non_perpetua['Conversion_Rate']) / non_perpetua['Conversion_Rate'] * 100) if non_perpetua['Conversion_Rate'] > 0 else 0
    }

    for metric, delta in deltas.items():
        print(f"{metric:20s}: {delta:+15,.2f}")

    print()
    print("Percentage Improvements:")
    for metric, pct in improvements.items():
        symbol = "✓" if pct > 0 else "✗"
        print(f"  {symbol} {metric:25s}: {pct:+7.2f}%")

    # Save summary
    summary = {
        'generated_at': datetime.now().isoformat(),
        'perpetua_asins': len(perpetua_asins),
        'non_perpetua_asins': len(non_perpetua_asins),
        'total_asins': len(all_asins),
        'campaign_records_analyzed': len(known_campaigns),
        'perpetua_performance': perpetua.to_dict(),
        'non_perpetua_performance': non_perpetua.to_dict(),
        'deltas': deltas,
        'improvements_pct': improvements
    }

    summary_file = AGG_DIR / 'perpetua_comparison_summary.json'
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2, default=str)

    print()
    print(f"✓ Saved summary to: {summary_file}")

# Save comparison table
comparison_file = AGG_DIR / 'perpetua_vs_non_perpetua.csv'
comparison.to_csv(comparison_file)
print(f"✓ Saved comparison table to: {comparison_file}")

print()
print("=" * 80)
print("✓ DATA PROCESSING COMPLETE")
print("=" * 80)
