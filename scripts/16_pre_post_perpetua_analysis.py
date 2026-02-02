#!/usr/bin/env python3
"""
PRE-PERPETUA vs POST-PERPETUA IMPLEMENTATION ANALYSIS
Before: Nov 15 - Dec 14, 2025 (Manual only)
After: Dec 15, 2025 onwards (Perpetua launched)
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import json

BASE_DIR = Path(__file__).parent.parent
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
OUTPUT_DIR = BASE_DIR / 'outputs'

print("=" * 100)
print("PRE-PERPETUA vs POST-PERPETUA IMPLEMENTATION ANALYSIS")
print("=" * 100)
print()

# ============================================================================
# LOAD DATA
# ============================================================================

print("[1/5] Loading data with Perpetua launch date context...")

# Load merged orders + advertising data
merged = pd.read_csv(PROCESSED_DIR / 'orders_advertising_merged.csv', low_memory=False)
merged['Date'] = pd.to_datetime(merged['Date'], errors='coerce')
merged = merged[merged['Date'].notna()]

# Define periods
PERPETUA_LAUNCH_DATE = pd.to_datetime('2025-12-15')
PRE_START = pd.to_datetime('2025-11-15')
PRE_END = pd.to_datetime('2025-12-14')

print(f"  ✓ Perpetua Launch Date: {PERPETUA_LAUNCH_DATE.date()}")
print(f"  ✓ Pre-Perpetua Period: {PRE_START.date()} to {PRE_END.date()} (30 days)")
print(f"  ✓ Post-Perpetua Period: {PERPETUA_LAUNCH_DATE.date()} onwards")

# Split data
pre_perpetua = merged[(merged['Date'] >= PRE_START) & (merged['Date'] <= PRE_END)]
post_perpetua = merged[merged['Date'] >= PERPETUA_LAUNCH_DATE]

print(f"\n  Pre-Perpetua records: {len(pre_perpetua):,}")
print(f"  Post-Perpetua records: {len(post_perpetua):,}")

# ============================================================================
# CALCULATE METRICS FOR BOTH PERIODS
# ============================================================================

print("\n[2/5] Calculating Pre vs Post metrics...")

def calc_period_metrics(df, period_name):
    """Calculate comprehensive metrics for a time period"""
    total_spend = df['Ad_Spend'].sum()
    ad_sales = df['Ad_Sales'].sum()
    total_revenue = df['Total_Revenue'].sum()
    organic_sales = df['Organic_Sales'].sum()

    # Count days in period
    days = df['Date'].nunique()

    return {
        'Period': period_name,
        'Days': days,
        'Total_Revenue': total_revenue,
        'Ad_Spend': total_spend,
        'Ad_Sales': ad_sales,
        'Organic_Sales': organic_sales,

        # Aggregate metrics (CORRECT method)
        'ROAS': ad_sales / total_spend if total_spend > 0 else 0,
        'ACOS': total_spend / ad_sales if ad_sales > 0 else 0,
        'TACoS': total_spend / total_revenue if total_revenue > 0 else 0,
        'T_ROAS': total_revenue / total_spend if total_spend > 0 else 0,
        'Organic_Ratio': organic_sales / total_revenue if total_revenue > 0 else 0,

        # Daily averages
        'Avg_Daily_Spend': total_spend / days if days > 0 else 0,
        'Avg_Daily_Revenue': total_revenue / days if days > 0 else 0,
        'Avg_Daily_Ad_Sales': ad_sales / days if days > 0 else 0,
        'Avg_Daily_Organic': organic_sales / days if days > 0 else 0,
    }

pre = calc_period_metrics(pre_perpetua, 'Pre-Perpetua (Manual)')
post = calc_period_metrics(post_perpetua, 'Post-Perpetua (SaaS)')

# ============================================================================
# DISPLAY COMPARISON
# ============================================================================

print(f"\n{'='*100}")
print("PRE-PERPETUA (Nov 15 - Dec 14, 2025) - MANUAL ADVERTISING ONLY")
print(f"{'='*100}")
print(f"  Period: {pre['Days']} days")
print(f"  Total Revenue: ${pre['Total_Revenue']:,.2f}")
print(f"  Ad Spend: ${pre['Ad_Spend']:,.2f}")
print(f"  Ad Sales: ${pre['Ad_Sales']:,.2f}")
print(f"  Organic Sales: ${pre['Organic_Sales']:,.2f}")
print(f"  ---")
print(f"  ROAS: {pre['ROAS']:.2f}x")
print(f"  ACOS: {pre['ACOS']*100:.1f}%")
print(f"  TACoS: {pre['TACoS']*100:.1f}%")
print(f"  T-ROAS: {pre['T_ROAS']:.2f}x")
print(f"  Organic Ratio: {pre['Organic_Ratio']*100:.1f}%")
print(f"  ---")
print(f"  Avg Daily Revenue: ${pre['Avg_Daily_Revenue']:,.2f}")
print(f"  Avg Daily Ad Spend: ${pre['Avg_Daily_Spend']:,.2f}")

print(f"\n{'='*100}")
print("POST-PERPETUA (Dec 15, 2025 onwards) - WITH PERPETUA SaaS")
print(f"{'='*100}")
print(f"  Period: {post['Days']} days")
print(f"  Total Revenue: ${post['Total_Revenue']:,.2f}")
print(f"  Ad Spend: ${post['Ad_Spend']:,.2f}")
print(f"  Ad Sales: ${post['Ad_Sales']:,.2f}")
print(f"  Organic Sales: ${post['Organic_Sales']:,.2f}")
print(f"  ---")
print(f"  ROAS: {post['ROAS']:.2f}x")
print(f"  ACOS: {post['ACOS']*100:.1f}%")
print(f"  TACoS: {post['TACoS']*100:.1f}%")
print(f"  T-ROAS: {post['T_ROAS']:.2f}x")
print(f"  Organic Ratio: {post['Organic_Ratio']*100:.1f}%")
print(f"  ---")
print(f"  Avg Daily Revenue: ${post['Avg_Daily_Revenue']:,.2f}")
print(f"  Avg Daily Ad Spend: ${post['Avg_Daily_Spend']:,.2f}")

# ============================================================================
# CALCULATE IMPACT
# ============================================================================

print(f"\n{'='*100}")
print("PERPETUA IMPLEMENTATION IMPACT")
print(f"{'='*100}\n")

impact = {
    'ROAS_Change': post['ROAS'] - pre['ROAS'],
    'ROAS_Change_Pct': ((post['ROAS'] - pre['ROAS']) / pre['ROAS'] * 100) if pre['ROAS'] > 0 else 0,
    'TACoS_Change': (post['TACoS'] - pre['TACoS']) * 100,  # Convert to percentage points
    'T_ROAS_Change': post['T_ROAS'] - pre['T_ROAS'],
    'Organic_Ratio_Change': (post['Organic_Ratio'] - pre['Organic_Ratio']) * 100,
    'Daily_Revenue_Change': post['Avg_Daily_Revenue'] - pre['Avg_Daily_Revenue'],
    'Daily_Revenue_Change_Pct': ((post['Avg_Daily_Revenue'] - pre['Avg_Daily_Revenue']) / pre['Avg_Daily_Revenue'] * 100) if pre['Avg_Daily_Revenue'] > 0 else 0,
}

print(f"{'Metric':<30} {'Change':>15} {'% Change':>12} {'Direction'}")
print("-" * 70)

metrics_to_show = [
    ('ROAS', impact['ROAS_Change'], impact['ROAS_Change_Pct'], 'Higher is better'),
    ('TACoS (pp)', impact['TACoS_Change'], None, 'Lower is better'),
    ('T-ROAS', impact['T_ROAS_Change'], None, 'Higher is better'),
    ('Organic Ratio (pp)', impact['Organic_Ratio_Change'], None, 'Higher is better'),
    ('Avg Daily Revenue', impact['Daily_Revenue_Change'], impact['Daily_Revenue_Change_Pct'], 'Higher is better'),
]

for metric, change, pct, direction in metrics_to_show:
    symbol = '✓' if change > 0 else '✗'
    if pct is not None:
        print(f"{metric:<30} {change:>+14.2f} {pct:>+11.1f}% {symbol} {direction}")
    else:
        print(f"{metric:<30} {change:>+14.2f} {'':>12} {symbol} {direction}")

# Annualized impact
if post['Days'] >= 30:
    annualized_revenue_impact = impact['Daily_Revenue_Change'] * 365
    print(f"\n💰 ANNUALIZED IMPACT:")
    print(f"   Daily revenue change: ${impact['Daily_Revenue_Change']:,.2f} ({impact['Daily_Revenue_Change_Pct']:+.1f}%)")
    print(f"   Annualized impact: ${annualized_revenue_impact:,.2f}")

# ============================================================================
# SAVE RESULTS
# ============================================================================

print("\n[3/5] Saving Pre/Post analysis...")

results = {
    'analysis_type': 'Pre-Perpetua vs Post-Perpetua Implementation',
    'perpetua_launch_date': PERPETUA_LAUNCH_DATE.isoformat(),
    'pre_period': {
        'start': PRE_START.isoformat(),
        'end': PRE_END.isoformat(),
        'metrics': pre
    },
    'post_period': {
        'start': PERPETUA_LAUNCH_DATE.isoformat(),
        'metrics': post
    },
    'impact': impact
}

with open(OUTPUT_DIR / 'pre_post_perpetua_analysis.json', 'w') as f:
    json.dump(results, f, indent=2, default=str)

print(f"  ✓ Saved to: {OUTPUT_DIR / 'pre_post_perpetua_analysis.json'}")

# ============================================================================
# KEY INSIGHTS
# ============================================================================

print(f"\n{'='*100}")
print("KEY INSIGHTS FROM PRE/POST ANALYSIS")
print(f"{'='*100}\n")

if impact['ROAS_Change'] > 0:
    print(f"✓ ROAS IMPROVED by {impact['ROAS_Change_Pct']:.0f}% after Perpetua launch")
    print(f"  Before: {pre['ROAS']:.2f}x → After: {post['ROAS']:.2f}x")
else:
    print(f"⚠ ROAS DECLINED by {abs(impact['ROAS_Change_Pct']):.0f}% after Perpetua launch")
    print(f"  Before: {pre['ROAS']:.2f}x → After: {post['ROAS']:.2f}x")

print()

if impact['TACoS_Change'] < 0:
    print(f"✓ TACoS IMPROVED (lower is better) by {abs(impact['TACoS_Change']):.1f} percentage points")
    print(f"  Before: {pre['TACoS']*100:.1f}% → After: {post['TACoS']*100:.1f}%")
else:
    print(f"⚠ TACoS INCREASED by {impact['TACoS_Change']:.1f} percentage points")
    print(f"  Before: {pre['TACoS']*100:.1f}% → After: {post['TACoS']*100:.1f}%")

print()

if impact['Daily_Revenue_Change'] > 0:
    print(f"✓ DAILY REVENUE INCREASED by ${impact['Daily_Revenue_Change']:,.0f} ({impact['Daily_Revenue_Change_Pct']:+.1f}%)")
    print(f"  Before: ${pre['Avg_Daily_Revenue']:,.0f}/day → After: ${post['Avg_Daily_Revenue']:,.0f}/day")
else:
    print(f"⚠ DAILY REVENUE DECREASED by ${abs(impact['Daily_Revenue_Change']):,.0f}")

print()
print("=" * 100)
print("✓ PRE/POST ANALYSIS COMPLETE")
print("=" * 100)
print("\nNext: Creating dashboard with Pre-Perpetua vs Post-Perpetua comparison")
