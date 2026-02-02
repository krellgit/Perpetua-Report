#!/usr/bin/env python3
"""
Year-over-Year Analysis + Ad Spend → Organic Sales Correlation
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
OUTPUT_DIR = BASE_DIR / 'outputs'

print("=" * 100)
print("YEAR-OVER-YEAR ANALYSIS + AD SPEND → ORGANIC SALES CORRELATION")
print("=" * 100)
print()

# ============================================================================
# LOAD CURRENT YEAR DATA
# ============================================================================

print("[1/4] Loading current year data (2025-2026)...")
merged = pd.read_csv(PROCESSED_DIR / 'orders_advertising_merged.csv', low_memory=False)
merged['Date'] = pd.to_datetime(merged['Date'], errors='coerce')
merged = merged[merged['Date'].notna()]

# Extract December 2025 and January 2026
merged['Month'] = merged['Date'].dt.to_period('M')

dec_2025 = merged[merged['Month'] == '2025-12']
jan_2026 = merged[merged['Month'] == '2026-01']

print(f"  ✓ December 2025: {len(dec_2025):,} records")
print(f"  ✓ January 2026: {len(jan_2026):,} records")

# Aggregate by month
def aggregate_month(df):
    return {
        'Ad_Spend': df['Ad_Spend'].sum(),
        'Ad_Sales': df['Ad_Sales'].sum(),
        'Total_Revenue': df['Total_Revenue'].sum(),
        'Organic_Sales': df['Organic_Sales'].sum(),
        'ROAS': df['Ad_Sales'].sum() / df['Ad_Spend'].sum() if df['Ad_Spend'].sum() > 0 else 0,
        'TACoS': (df['Ad_Spend'].sum() / df['Total_Revenue'].sum() * 100) if df['Total_Revenue'].sum() > 0 else 0,
        'T_ROAS': df['Total_Revenue'].sum() / df['Ad_Spend'].sum() if df['Ad_Spend'].sum() > 0 else 0
    }

dec_2025_agg = aggregate_month(dec_2025)
jan_2026_agg = aggregate_month(jan_2026)

print(f"\nDecember 2025:")
print(f"  Ad Spend: ${dec_2025_agg['Ad_Spend']:,.0f}")
print(f"  Ad Sales: ${dec_2025_agg['Ad_Sales']:,.0f}")
print(f"  Total Revenue: ${dec_2025_agg['Total_Revenue']:,.0f}")
print(f"  ROAS: {dec_2025_agg['ROAS']:.2f}x")
print(f"  TACoS: {dec_2025_agg['TACoS']:.1f}%")

print(f"\nJanuary 2026:")
print(f"  Ad Spend: ${jan_2026_agg['Ad_Spend']:,.0f}")
print(f"  Ad Sales: ${jan_2026_agg['Ad_Sales']:,.0f}")
print(f"  Total Revenue: ${jan_2026_agg['Total_Revenue']:,.0f}")
print(f"  ROAS: {jan_2026_agg['ROAS']:.2f}x")
print(f"  TACoS: {jan_2026_agg['TACoS']:.1f}%")

# ============================================================================
# YEAR-OVER-YEAR COMPARISON
# ============================================================================

print(f"\n{'='*100}")
print("YEAR-OVER-YEAR COMPARISON")
print(f"{'='*100}\n")

# Last year data (provided by user)
dec_2024 = {
    'Ad_Sales': 91526,
    'Ad_Spend': 108688,
    'ROAS': 91526 / 108688
}

jan_2024 = {
    'Ad_Sales': 174853,
    'Ad_Spend': 139723,
    'ROAS': 174853 / 139723
}

# December YoY
print("DECEMBER COMPARISON (2024 vs 2025):")
print(f"{'='*60}")
print(f"{'Metric':<25} {'2024':>15} {'2025':>15} {'Change':>15} {'%':>10}")
print(f"{'-'*60}")
print(f"{'Ad Spend':<25} ${dec_2024['Ad_Spend']:>14,} ${dec_2025_agg['Ad_Spend']:>14,.0f} ${dec_2025_agg['Ad_Spend']-dec_2024['Ad_Spend']:>14,.0f} {((dec_2025_agg['Ad_Spend']-dec_2024['Ad_Spend'])/dec_2024['Ad_Spend']*100):>9.1f}%")
print(f"{'Ad Sales':<25} ${dec_2024['Ad_Sales']:>14,} ${dec_2025_agg['Ad_Sales']:>14,.0f} ${dec_2025_agg['Ad_Sales']-dec_2024['Ad_Sales']:>14,.0f} {((dec_2025_agg['Ad_Sales']-dec_2024['Ad_Sales'])/dec_2024['Ad_Sales']*100):>9.1f}%")
print(f"{'ROAS':<25} {dec_2024['ROAS']:>14.2f}x {dec_2025_agg['ROAS']:>14.2f}x {dec_2025_agg['ROAS']-dec_2024['ROAS']:>14.2f}x {((dec_2025_agg['ROAS']-dec_2024['ROAS'])/dec_2024['ROAS']*100):>9.1f}%")

# January YoY
print(f"\nJANUARY COMPARISON (2024 vs 2026):")
print(f"{'='*60}")
print(f"{'Metric':<25} {'2024':>15} {'2026':>15} {'Change':>15} {'%':>10}")
print(f"{'-'*60}")
print(f"{'Ad Spend':<25} ${jan_2024['Ad_Spend']:>14,} ${jan_2026_agg['Ad_Spend']:>14,.0f} ${jan_2026_agg['Ad_Spend']-jan_2024['Ad_Spend']:>14,.0f} {((jan_2026_agg['Ad_Spend']-jan_2024['Ad_Spend'])/jan_2024['Ad_Spend']*100):>9.1f}%")
print(f"{'Ad Sales':<25} ${jan_2024['Ad_Sales']:>14,} ${jan_2026_agg['Ad_Sales']:>14,.0f} ${jan_2026_agg['Ad_Sales']-jan_2024['Ad_Sales']:>14,.0f} {((jan_2026_agg['Ad_Sales']-jan_2024['Ad_Sales'])/jan_2024['Ad_Sales']*100):>9.1f}%")
print(f"{'ROAS':<25} {jan_2024['ROAS']:>14.2f}x {jan_2026_agg['ROAS']:>14.2f}x {jan_2026_agg['ROAS']-jan_2024['ROAS']:>14.2f}x {((jan_2026_agg['ROAS']-jan_2024['ROAS'])/jan_2024['ROAS']*100):>9.1f}%")

# Key insights
print(f"\n{'='*60}")
print("KEY YoY INSIGHTS:")
print(f"{'='*60}\n")

dec_roas_improvement = ((dec_2025_agg['ROAS'] - dec_2024['ROAS']) / dec_2024['ROAS'] * 100)
jan_roas_improvement = ((jan_2026_agg['ROAS'] - jan_2024['ROAS']) / jan_2024['ROAS'] * 100)

if dec_roas_improvement > 0:
    print(f"✓ December ROAS improved {dec_roas_improvement:.0f}% YoY ({dec_2024['ROAS']:.2f}x → {dec_2025_agg['ROAS']:.2f}x)")
else:
    print(f"⚠ December ROAS declined {abs(dec_roas_improvement):.0f}% YoY ({dec_2024['ROAS']:.2f}x → {dec_2025_agg['ROAS']:.2f}x)")

if jan_roas_improvement > 0:
    print(f"✓ January ROAS improved {jan_roas_improvement:.0f}% YoY ({jan_2024['ROAS']:.2f}x → {jan_2026_agg['ROAS']:.2f}x)")
else:
    print(f"⚠ January ROAS declined {abs(jan_roas_improvement):.0f}% YoY ({jan_2024['ROAS']:.2f}x → {jan_2026_agg['ROAS']:.2f}x)")

# ============================================================================
# AD SPEND → ORGANIC SALES CORRELATION ANALYSIS
# ============================================================================

print(f"\n{'='*100}")
print("AD SPEND → ORGANIC SALES CORRELATION ANALYSIS")
print(f"{'='*100}\n")

# Daily data by platform
perpetua_daily = merged[merged['Advertising_Type'] == 'Perpetua'].groupby('Date').agg({
    'Ad_Spend': 'sum',
    'Organic_Sales': 'sum',
    'Total_Revenue': 'sum'
}).reset_index().sort_values('Date')

non_perpetua_daily = merged[merged['Advertising_Type'] == 'Non-Perpetua'].groupby('Date').agg({
    'Ad_Spend': 'sum',
    'Organic_Sales': 'sum',
    'Total_Revenue': 'sum'
}).reset_index().sort_values('Date')

print("Testing correlation at different time lags...")
print()

# Test correlation with lags (0, 7, 14, 30 days)
lags = [0, 7, 14, 30]

print(f"{'Platform':<15} {'Lag (days)':<12} {'Correlation':<15} {'P-Value':<12} {'Significant?'}")
print("-" * 70)

for platform_name, daily_df in [('Perpetua', perpetua_daily), ('Non-Perpetua', non_perpetua_daily)]:
    for lag in lags:
        if len(daily_df) > lag:
            # Shift organic sales by lag days
            daily_df[f'Organic_Lag_{lag}'] = daily_df['Organic_Sales'].shift(-lag)

            # Get complete pairs (no NaN)
            valid = daily_df[['Ad_Spend', f'Organic_Lag_{lag}']].dropna()

            if len(valid) > 10:  # Need sufficient data points
                try:
                    # Use numpy correlation
                    corr = np.corrcoef(valid['Ad_Spend'], valid[f'Organic_Lag_{lag}'])[0, 1]
                    # Approximate p-value (rough estimate)
                    n = len(valid)
                    t_stat = corr * np.sqrt(n - 2) / np.sqrt(1 - corr**2) if abs(corr) < 1 else 0
                    p_value = 0.01 if abs(t_stat) > 2.5 else 0.05 if abs(t_stat) > 2 else 0.5
                    sig = "YES" if p_value < 0.05 else "NO"
                    print(f"{platform_name:<15} {lag:<12} {corr:>14.3f} {p_value:>11.4f}  {sig}")
                except:
                    print(f"{platform_name:<15} {lag:<12} {'N/A':<15} {'N/A':<12} N/A")

print()
print("INTERPRETATION:")
print("  • Positive correlation = Higher ad spend associated with higher organic sales")
print("  • Negative correlation = No relationship or inverse relationship")
print("  • P-value < 0.05 = Statistically significant")
print("  • Lag shows delay: 7-day lag = ad spend impact shows 7 days later")

# Calculate elasticity (% change in organic per % change in ad spend)
print(f"\n{'='*100}")
print("AD SPEND ELASTICITY (% Organic Change per 1% Ad Spend Change)")
print(f"{'='*100}\n")

for platform_name, daily_df in [('Perpetua', perpetua_daily), ('Non-Perpetua', non_perpetua_daily)]:
    if len(daily_df) > 2:
        # Calculate percentage changes
        daily_df['Ad_Spend_Pct_Change'] = daily_df['Ad_Spend'].pct_change()
        daily_df['Organic_Pct_Change'] = daily_df['Organic_Sales'].pct_change()

        # Remove infinites and NaN
        valid = daily_df[['Ad_Spend_Pct_Change', 'Organic_Pct_Change']].replace([np.inf, -np.inf], np.nan).dropna()

        if len(valid) > 10:
            # Regression: Organic % Change = f(Ad Spend % Change)
            from numpy.polynomial import polynomial as P
            coeffs = np.polyfit(valid['Ad_Spend_Pct_Change'], valid['Organic_Pct_Change'], 1)
            slope = coeffs[0]

            print(f"{platform_name}:")
            print(f"  Elasticity: {slope:.2f}")
            print(f"  Interpretation: 1% increase in ad spend → {slope:.2f}% change in organic sales")
            print()

# Save YoY comparison
yoy_data = {
    'december': {
        '2024': dec_2024,
        '2025': {
            'Ad_Sales': dec_2025_agg['Ad_Sales'],
            'Ad_Spend': dec_2025_agg['Ad_Spend'],
            'ROAS': dec_2025_agg['ROAS'],
            'Total_Revenue': dec_2025_agg['Total_Revenue'],
            'TACoS': dec_2025_agg['TACoS']
        },
        'growth': {
            'Ad_Sales': ((dec_2025_agg['Ad_Sales'] - dec_2024['Ad_Sales']) / dec_2024['Ad_Sales'] * 100),
            'Ad_Spend': ((dec_2025_agg['Ad_Spend'] - dec_2024['Ad_Spend']) / dec_2024['Ad_Spend'] * 100),
            'ROAS': ((dec_2025_agg['ROAS'] - dec_2024['ROAS']) / dec_2024['ROAS'] * 100)
        }
    },
    'january': {
        '2024': jan_2024,
        '2026': {
            'Ad_Sales': jan_2026_agg['Ad_Sales'],
            'Ad_Spend': jan_2026_agg['Ad_Spend'],
            'ROAS': jan_2026_agg['ROAS'],
            'Total_Revenue': jan_2026_agg['Total_Revenue'],
            'TACoS': jan_2026_agg['TACoS']
        },
        'growth': {
            'Ad_Sales': ((jan_2026_agg['Ad_Sales'] - jan_2024['Ad_Sales']) / jan_2024['Ad_Sales'] * 100),
            'Ad_Spend': ((jan_2026_agg['Ad_Spend'] - jan_2024['Ad_Spend']) / jan_2024['Ad_Spend'] * 100),
            'ROAS': ((jan_2026_agg['ROAS'] - jan_2024['ROAS']) / jan_2024['ROAS'] * 100)
        }
    }
}

import json
with open(OUTPUT_DIR / 'yoy_analysis.json', 'w') as f:
    json.dump(yoy_data, f, indent=2, default=str)

print("\n" + "="*100)
print("✓ ANALYSIS COMPLETE")
print("="*100)
print(f"\nYoY comparison saved to: {OUTPUT_DIR / 'yoy_analysis.json'}")
print("\nKey Findings:")
print(f"  Dec: {dec_roas_improvement:+.0f}% ROAS change YoY")
print(f"  Jan: {jan_roas_improvement:+.0f}% ROAS change YoY")
print("\nCorrelation results show if ad spend drives organic sales")
