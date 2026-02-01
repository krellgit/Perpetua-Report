#!/usr/bin/env python3
"""
ASIN-Level Performance Analysis: Perpetua vs Non-Perpetua
Uses Advertised Products report for comprehensive ASIN-level metrics
"""

import pandas as pd
import json
from pathlib import Path
from datetime import datetime

# Paths
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / 'data' / 'recent-reports'
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
AGG_DIR = BASE_DIR / 'data' / 'aggregated'

print("=" * 80)
print("ASIN-LEVEL PERFORMANCE ANALYSIS: PERPETUA VS NON-PERPETUA")
print("=" * 80)
print()

# Step 1: Load ASIN lists
print("[1/5] Loading ASIN lists...")
perpetua_df = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='perpetua list')
perpetua_asins = set(perpetua_df['ASIN'].dropna().str.strip().tolist())
print(f"  ✓ Perpetua ASINs: {len(perpetua_asins)}")

all_asins_df = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='All ASIns')
all_asins = set(all_asins_df['ASIN (Informational only)'].dropna().str.strip().tolist())
non_perpetua_asins = all_asins - perpetua_asins
print(f"  ✓ All ASINs: {len(all_asins)}")
print(f"  ✓ Non-Perpetua ASINs: {len(non_perpetua_asins)}")
print()

# Step 2: Load Advertised Products report
print("[2/5] Loading Advertised Products report...")
print("  (This may take a moment - 14MB file)")
ad_products = pd.read_excel(DATA_DIR / 'SP_Advertised_Products_-_Max (1).xlsx')
print(f"  ✓ Loaded {len(ad_products):,} product advertising records")
print(f"  ✓ Date range: {ad_products['Date'].min()} to {ad_products['Date'].max()}")
print(f"  ✓ Unique ASINs: {ad_products['Advertised ASIN'].nunique()}")
print()

# Step 3: Tag ASINs
print("[3/5] Classifying ASINs as Perpetua vs Non-Perpetua...")

def classify_asin(asin):
    if pd.isna(asin):
        return 'Unknown'
    asin = str(asin).strip()
    if asin in perpetua_asins:
        return 'Perpetua'
    elif asin in non_perpetua_asins:
        return 'Non-Perpetua'
    else:
        return 'Unknown'

ad_products['Advertising_Type'] = ad_products['Advertised ASIN'].apply(classify_asin)

# Distribution
type_counts = ad_products['Advertising_Type'].value_counts()
print(f"  ASIN Classification:")
for ad_type, count in type_counts.items():
    pct = count / len(ad_products) * 100
    print(f"    {ad_type:15s}: {count:8,} records ({pct:5.1f}%)")

# Count unique ASINs by type
unique_asins = ad_products.groupby('Advertising_Type')['Advertised ASIN'].nunique()
print(f"\n  Unique ASINs by Type:")
for ad_type, count in unique_asins.items():
    print(f"    {ad_type:15s}: {count:4} ASINs")
print()

# Step 4: Clean and aggregate data
print("[4/5] Aggregating ASIN-level metrics...")

# Filter to known ASINs
known = ad_products[ad_products['Advertising_Type'].isin(['Perpetua', 'Non-Perpetua'])].copy()
print(f"  ✓ Analyzing {len(known):,} records from {known['Advertised ASIN'].nunique()} ASINs")

# Clean numeric columns
numeric_cols = ['Spend', 'Cost Per Click (CPC)', 'Impressions', 'Clicks',
                '7 Day Total Orders (#)', '7 Day Total Sales ', '7 Day Total Units (#)']

for col in numeric_cols:
    if col in known.columns:
        if known[col].dtype == 'object':
            known[col] = known[col].astype(str).str.replace('$', '').str.replace('%', '').str.replace(',', '')
        known[col] = pd.to_numeric(known[col], errors='coerce').fillna(0)

# Calculate metrics
if 'Total Advertising Cost of Sales (ACOS) ' in known.columns:
    known['ACOS'] = pd.to_numeric(
        known['Total Advertising Cost of Sales (ACOS) '].astype(str).str.replace('%', ''),
        errors='coerce'
    ).fillna(0) / 100
else:
    known['ACOS'] = known.apply(
        lambda x: x['Spend'] / x['7 Day Total Sales '] if x['7 Day Total Sales '] > 0 else 0,
        axis=1
    )

if 'Total Return on Advertising Spend (ROAS)' in known.columns:
    known['ROAS'] = pd.to_numeric(known['Total Return on Advertising Spend (ROAS)'], errors='coerce').fillna(0)
else:
    known['ROAS'] = known.apply(
        lambda x: x['7 Day Total Sales '] / x['Spend'] if x['Spend'] > 0 else 0,
        axis=1
    )

if '7 Day Conversion Rate' in known.columns:
    known['Conversion_Rate'] = pd.to_numeric(
        known['7 Day Conversion Rate'].astype(str).str.replace('%', ''),
        errors='coerce'
    ).fillna(0) / 100
else:
    known['Conversion_Rate'] = known.apply(
        lambda x: x['7 Day Total Orders (#)'] / x['Clicks'] if x['Clicks'] > 0 else 0,
        axis=1
    )

# Aggregate by Advertising Type
comparison = known.groupby('Advertising_Type').agg({
    'Spend': 'sum',
    '7 Day Total Sales ': 'sum',
    '7 Day Total Orders (#)': 'sum',
    '7 Day Total Units (#)': 'sum',
    'Impressions': 'sum',
    'Clicks': 'sum',
    'Advertised ASIN': 'nunique'
}).round(2)

comparison.columns = ['Total_Spend', 'Total_Sales', 'Total_Orders', 'Total_Units',
                      'Total_Impressions', 'Total_Clicks', 'Unique_ASINs']

# Calculate average metrics
avg_metrics = known.groupby('Advertising_Type').agg({
    'ACOS': 'mean',
    'ROAS': 'mean',
    'Cost Per Click (CPC)': 'mean',
    'Conversion_Rate': 'mean'
}).round(4)

# Combine
comparison = comparison.join(avg_metrics)

# Calculate derived metrics
comparison['Avg_CPC'] = comparison['Total_Spend'] / comparison['Total_Clicks']
comparison['Avg_CVR'] = comparison['Total_Orders'] / comparison['Total_Clicks']
comparison['CTR'] = comparison['Total_Clicks'] / comparison['Total_Impressions']

# Calculate per-ASIN averages
comparison['Spend_Per_ASIN'] = comparison['Total_Spend'] / comparison['Unique_ASINs']
comparison['Sales_Per_ASIN'] = comparison['Total_Sales'] / comparison['Unique_ASINs']
comparison['Orders_Per_ASIN'] = comparison['Total_Orders'] / comparison['Unique_ASINs']

print()
print("=" * 80)
print("PERPETUA VS NON-PERPETUA: AGGREGATE METRICS")
print("=" * 80)
print()
print(comparison.to_string())
print()

# Step 5: Calculate deltas and improvements
print("=" * 80)
print("PERFORMANCE COMPARISON & INSIGHTS")
print("=" * 80)
print()

if 'Perpetua' in comparison.index and 'Non-Perpetua' in comparison.index:
    perpetua = comparison.loc['Perpetua']
    non_perpetua = comparison.loc['Non-Perpetua']

    # Calculate improvements
    metrics_comparison = {
        'Total Spend': {
            'Perpetua': perpetua['Total_Spend'],
            'Non-Perpetua': non_perpetua['Total_Spend'],
            'Delta': perpetua['Total_Spend'] - non_perpetua['Total_Spend'],
            'Better': 'Higher is worse' if perpetua['Total_Spend'] > non_perpetua['Total_Spend'] else 'Lower is better'
        },
        'Total Sales': {
            'Perpetua': perpetua['Total_Sales'],
            'Non-Perpetua': non_perpetua['Total_Sales'],
            'Delta': perpetua['Total_Sales'] - non_perpetua['Total_Sales'],
            'Pct_Change': ((perpetua['Total_Sales'] - non_perpetua['Total_Sales']) / non_perpetua['Total_Sales'] * 100) if non_perpetua['Total_Sales'] > 0 else 0
        },
        'Avg ACOS': {
            'Perpetua': perpetua['ACOS'],
            'Non-Perpetua': non_perpetua['ACOS'],
            'Delta': perpetua['ACOS'] - non_perpetua['ACOS'],
            'Improvement_Pct': ((non_perpetua['ACOS'] - perpetua['ACOS']) / non_perpetua['ACOS'] * 100) if non_perpetua['ACOS'] > 0 else 0
        },
        'Avg ROAS': {
            'Perpetua': perpetua['ROAS'],
            'Non-Perpetua': non_perpetua['ROAS'],
            'Delta': perpetua['ROAS'] - non_perpetua['ROAS'],
            'Improvement_Pct': ((perpetua['ROAS'] - non_perpetua['ROAS']) / non_perpetua['ROAS'] * 100) if non_perpetua['ROAS'] > 0 else 0
        },
        'Avg CPC': {
            'Perpetua': perpetua['Avg_CPC'],
            'Non-Perpetua': non_perpetua['Avg_CPC'],
            'Delta': perpetua['Avg_CPC'] - non_perpetua['Avg_CPC'],
            'Improvement_Pct': ((non_perpetua['Avg_CPC'] - perpetua['Avg_CPC']) / non_perpetua['Avg_CPC'] * 100) if non_perpetua['Avg_CPC'] > 0 else 0
        },
        'Conversion Rate': {
            'Perpetua': perpetua['Avg_CVR'],
            'Non-Perpetua': non_perpetua['Avg_CVR'],
            'Delta': perpetua['Avg_CVR'] - non_perpetua['Avg_CVR'],
            'Improvement_Pct': ((perpetua['Avg_CVR'] - non_perpetua['Avg_CVR']) / non_perpetua['Avg_CVR'] * 100) if non_perpetua['Avg_CVR'] > 0 else 0
        }
    }

    # Print comparison table
    print(f"{'Metric':<20} | {'Perpetua':>12} | {'Non-Perpetua':>12} | {'Delta':>12} | {'% Change':>10}")
    print("-" * 80)

    for metric, data in metrics_comparison.items():
        perpetua_val = data['Perpetua']
        non_perpetua_val = data['Non-Perpetua']
        delta = data['Delta']

        if 'Improvement_Pct' in data:
            pct = data['Improvement_Pct']
            symbol = "✓" if pct > 0 else ("=" if pct == 0 else "✗")
            print(f"{metric:<20} | {perpetua_val:>12.2f} | {non_perpetua_val:>12.2f} | {delta:>+12.2f} | {symbol} {pct:>+7.1f}%")
        elif 'Pct_Change' in data:
            pct = data['Pct_Change']
            print(f"{metric:<20} | ${perpetua_val:>11,.0f} | ${non_perpetua_val:>11,.0f} | ${delta:>+11,.0f} | {pct:>+9.1f}%")
        else:
            print(f"{metric:<20} | ${perpetua_val:>11,.0f} | ${non_perpetua_val:>11,.0f} | ${delta:>+11,.0f} | ")

    print()
    print("KEY INSIGHTS:")
    print()

    # ACOS insight
    if metrics_comparison['Avg ACOS']['Improvement_Pct'] > 0:
        print(f"  ✓ Perpetua ASINs have {metrics_comparison['Avg ACOS']['Improvement_Pct']:.1f}% BETTER ACOS (lower is better)")
        print(f"    Perpetua: {perpetua['ACOS']*100:.1f}% vs Non-Perpetua: {non_perpetua['ACOS']*100:.1f}%")
    else:
        print(f"  ✗ Perpetua ASINs have {abs(metrics_comparison['Avg ACOS']['Improvement_Pct']):.1f}% WORSE ACOS")
        print(f"    Perpetua: {perpetua['ACOS']*100:.1f}% vs Non-Perpetua: {non_perpetua['ACOS']*100:.1f}%")

    print()

    # ROAS insight
    if metrics_comparison['Avg ROAS']['Improvement_Pct'] > 0:
        print(f"  ✓ Perpetua ASINs have {metrics_comparison['Avg ROAS']['Improvement_Pct']:.1f}% BETTER ROAS (higher is better)")
        print(f"    Perpetua: {perpetua['ROAS']:.2f} vs Non-Perpetua: {non_perpetua['ROAS']:.2f}")
    else:
        print(f"  ✗ Perpetua ASINs have {abs(metrics_comparison['Avg ROAS']['Improvement_Pct']):.1f}% WORSE ROAS")
        print(f"    Perpetua: {perpetua['ROAS']:.2f} vs Non-Perpetua: {non_perpetua['ROAS']:.2f}")

    print()

    # Efficiency insight
    if metrics_comparison['Avg CPC']['Improvement_Pct'] > 0:
        print(f"  ✓ Perpetua has {metrics_comparison['Avg CPC']['Improvement_Pct']:.1f}% LOWER CPC (cost efficiency)")
        print(f"    Perpetua: ${perpetua['Avg_CPC']:.2f} vs Non-Perpetua: ${non_perpetua['Avg_CPC']:.2f}")
    else:
        print(f"  ⚠ Perpetua has {abs(metrics_comparison['Avg CPC']['Improvement_Pct']):.1f}% HIGHER CPC")
        print(f"    Perpetua: ${perpetua['Avg_CPC']:.2f} vs Non-Perpetua: ${non_perpetua['Avg_CPC']:.2f}")

    print()

    # Conversion insight
    if metrics_comparison['Conversion Rate']['Improvement_Pct'] > 0:
        print(f"  ✓ Perpetua has {metrics_comparison['Conversion Rate']['Improvement_Pct']:.1f}% BETTER conversion rate")
        print(f"    Perpetua: {perpetua['Avg_CVR']*100:.2f}% vs Non-Perpetua: {non_perpetua['Avg_CVR']*100:.2f}%")
    else:
        print(f"  ⚠ Perpetua has {abs(metrics_comparison['Conversion Rate']['Improvement_Pct']):.1f}% LOWER conversion rate")
        print(f"    Perpetua: {perpetua['Avg_CVR']*100:.2f}% vs Non-Perpetua: {non_perpetua['Avg_CVR']*100:.2f}%")

    # Save results
    summary = {
        'generated_at': datetime.now().isoformat(),
        'perpetua_asins_count': len(perpetua_asins),
        'non_perpetua_asins_count': len(non_perpetua_asins),
        'perpetua_asins_analyzed': int(perpetua['Unique_ASINs']),
        'non_perpetua_asins_analyzed': int(non_perpetua['Unique_ASINs']),
        'perpetua_metrics': perpetua.to_dict(),
        'non_perpetua_metrics': non_perpetua.to_dict(),
        'comparison': metrics_comparison
    }

    summary_file = AGG_DIR / 'asin_level_comparison.json'
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2, default=str)

    print()
    print(f"✓ Saved summary to: {summary_file}")

elif 'Perpetua' in comparison.index:
    print("  ⚠ Only Perpetua data available - no non-Perpetua campaigns found")
    print("    This suggests all advertised products are managed through Perpetua")
else:
    print("  ⚠ No Perpetua campaigns found in this dataset")

# Save full comparison
comparison_file = AGG_DIR / 'asin_comparison_full.csv'
comparison.to_csv(comparison_file)
print(f"✓ Saved comparison table to: {comparison_file}")

# Save processed data
processed_file = PROCESSED_DIR / 'advertised_products_processed.csv'
known.to_csv(processed_file, index=False)
print(f"✓ Saved processed data to: {processed_file}")

print()
print("=" * 80)
print("✓ ASIN-LEVEL ANALYSIS COMPLETE")
print("=" * 80)
