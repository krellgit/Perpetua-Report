#!/usr/bin/env python3
"""
Process Order Data and Calculate TACoS/T-ROAS
Integrates order reports with advertising data for comprehensive analysis
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / 'data' / 'recent-reports'
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
OUTPUT_DIR = BASE_DIR / 'outputs'

print("=" * 100)
print("ORDER DATA PROCESSING FOR TACoS AND T-ROAS ANALYSIS")
print("=" * 100)
print()

# ============================================================================
# STEP 1: LOAD ORDER DATA FILES
# ============================================================================

print("[1/8] Loading order data files...")
print("  (This may take a moment - 146 MB of data)")

# Column names from Amazon Order Report structure
order_columns = [
    'amazon-order-id', 'merchant-order-id', 'purchase-date', 'last-updated-date',
    'order-status', 'fulfillment-channel', 'sales-channel', 'order-channel',
    'url', 'ship-service-level', 'product-name', 'sku', 'asin', 'item-status',
    'quantity', 'currency', 'item-price', 'item-tax', 'shipping-price',
    'shipping-tax', 'gift-wrap-price', 'gift-wrap-tax', 'item-promotion-discount',
    'ship-promotion-discount', 'ship-city', 'ship-state', 'ship-postal-code',
    'ship-country', 'promotion-ids', 'is-business-order', 'purchase-order-number',
    'price-designation', 'fulfilled-by', 'is-iba', 'signature-confirmation-recommended',
    'buyer-name'
]

# Load both order files
orders1 = pd.read_csv(DATA_DIR / '212008020460 (1).txt', sep='\t',
                      names=order_columns, header=0, low_memory=False)
print(f"  ✓ File 1: {len(orders1):,} order lines (Dec 1 - Jan 1)")

orders2 = pd.read_csv(DATA_DIR / '215564020486.txt', sep='\t',
                      names=order_columns, header=0, low_memory=False)
print(f"  ✓ File 2: {len(orders2):,} order lines (Jan 1 - Feb 1)")

# Combine
orders = pd.concat([orders1, orders2], ignore_index=True)
print(f"  ✓ Combined: {len(orders):,} total order lines")

# ============================================================================
# STEP 2: CLEAN AND DE-DUPLICATE
# ============================================================================

print("[2/8] Cleaning and de-duplicating order data...")

# Filter to Shipped orders only
orders_clean = orders[orders['order-status'] == 'Shipped'].copy()
print(f"  ✓ Shipped orders: {len(orders_clean):,} ({len(orders_clean)/len(orders)*100:.1f}%)")

# Remove duplicates
before_dedup = len(orders_clean)
orders_clean = orders_clean.drop_duplicates(subset=['amazon-order-id', 'sku', 'quantity'])
print(f"  ✓ De-duplicated: {before_dedup:,} → {len(orders_clean):,} ({before_dedup - len(orders_clean):,} duplicates removed)")

# Convert dates
orders_clean['purchase-date'] = pd.to_datetime(orders_clean['purchase-date'], errors='coerce')
orders_clean = orders_clean[orders_clean['purchase-date'].notna()]

# Convert item-price to numeric
orders_clean['item-price'] = pd.to_numeric(orders_clean['item-price'], errors='coerce').fillna(0)
orders_clean['quantity'] = pd.to_numeric(orders_clean['quantity'], errors='coerce').fillna(0)

# Filter to valid prices
orders_clean = orders_clean[orders_clean['item-price'] > 0]
print(f"  ✓ Valid prices: {len(orders_clean):,} order lines")

# Calculate total revenue per line
orders_clean['revenue'] = orders_clean['item-price'] * orders_clean['quantity']

# Date range
min_order_date = orders_clean['purchase-date'].min()
max_order_date = orders_clean['purchase-date'].max()
print(f"  ✓ Order date range: {min_order_date.date()} to {max_order_date.date()}")

# ============================================================================
# STEP 3: AGGREGATE ORDERS BY DATE + SKU
# ============================================================================

print("[3/8] Aggregating orders by Date and SKU...")

# Extract just the date (no time)
orders_clean['Date'] = orders_clean['purchase-date'].dt.date
orders_clean['Date'] = pd.to_datetime(orders_clean['Date'])

# Aggregate
order_summary = orders_clean.groupby(['Date', 'sku']).agg({
    'revenue': 'sum',
    'quantity': 'sum',
    'amazon-order-id': 'nunique'
}).reset_index()

order_summary.columns = ['Date', 'SKU', 'Total_Revenue', 'Total_Units', 'Order_Count']

print(f"  ✓ Aggregated to {len(order_summary):,} Date+SKU combinations")
print(f"  ✓ Unique SKUs in orders: {order_summary['SKU'].nunique()}")
print(f"  ✓ Total revenue: ${order_summary['Total_Revenue'].sum():,.2f}")

# ============================================================================
# STEP 4: LOAD PERPETUA/NON-PERPETUA MAPPING
# ============================================================================

print("[4/8] Loading Perpetua SKU mappings...")

perpetua_list = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='perpetua list')
perpetua_skus = set(perpetua_list['SKU'].dropna().str.strip())

all_asins_list = pd.read_excel(DATA_DIR / 'ASIN list - perpetua.xlsx', sheet_name='All ASIns')
all_skus = set(all_asins_list['SKU'].dropna().str.strip())

non_perpetua_skus = all_skus - perpetua_skus

print(f"  ✓ Perpetua SKUs: {len(perpetua_skus)}")
print(f"  ✓ Non-Perpetua SKUs: {len(non_perpetua_skus)}")

# Tag orders
order_summary['Advertising_Type'] = order_summary['SKU'].apply(
    lambda x: 'Perpetua' if x in perpetua_skus else ('Non-Perpetua' if x in non_perpetua_skus else 'Unknown')
)

type_counts = order_summary['Advertising_Type'].value_counts()
print(f"\n  Order Classification:")
for ad_type, count in type_counts.items():
    revenue = order_summary[order_summary['Advertising_Type'] == ad_type]['Total_Revenue'].sum()
    print(f"    {ad_type:15s}: {count:6,} records, ${revenue:,.0f} revenue")

# ============================================================================
# STEP 5: LOAD ADVERTISING DATA
# ============================================================================

print("\n[5/8] Loading advertising data for TACoS calculation...")

# Load processed advertising data
ad_data = pd.read_csv(PROCESSED_DIR / 'advertised_products_processed.csv', low_memory=False)
ad_data['Date'] = pd.to_datetime(ad_data['Date'], errors='coerce')
ad_data = ad_data[ad_data['Date'].notna()]

# Clean numerics
for col in ['Spend', '7 Day Total Sales ']:
    ad_data[col] = pd.to_numeric(ad_data[col], errors='coerce').fillna(0)

# Aggregate ad data by Date + SKU
ad_summary = ad_data.groupby(['Date', 'Advertised SKU', 'Advertising_Type']).agg({
    'Spend': 'sum',
    '7 Day Total Sales ': 'sum'
}).reset_index()

ad_summary.columns = ['Date', 'SKU', 'Advertising_Type_Ad', 'Ad_Spend', 'Ad_Sales']

print(f"  ✓ Ad data: {len(ad_summary):,} Date+SKU combinations")

# ============================================================================
# STEP 6: MERGE ORDERS WITH ADVERTISING
# ============================================================================

print("[6/8] Merging order data with advertising data...")

# Merge on Date + SKU
merged = pd.merge(
    order_summary[order_summary['Advertising_Type'].isin(['Perpetua', 'Non-Perpetua'])],
    ad_summary,
    on=['Date', 'SKU'],
    how='outer',
    indicator=True
)

print(f"  ✓ Merged dataset: {len(merged):,} records")
print(f"  Merge breakdown:")
print(f"    Both (orders + ads): {len(merged[merged['_merge'] == 'both']):,}")
print(f"    Orders only: {len(merged[merged['_merge'] == 'left_only']):,}")
print(f"    Ads only: {len(merged[merged['_merge'] == 'right_only']):,}")

# Fill NaN values
merged['Total_Revenue'] = merged['Total_Revenue'].fillna(0)
merged['Ad_Spend'] = merged['Ad_Spend'].fillna(0)
merged['Ad_Sales'] = merged['Ad_Sales'].fillna(0)
merged['Advertising_Type'] = merged['Advertising_Type'].fillna(merged['Advertising_Type_Ad'])

# Calculate TACoS and T-ROAS
merged['Organic_Sales'] = merged['Total_Revenue'] - merged['Ad_Sales']
merged['TACoS'] = np.where(merged['Total_Revenue'] > 0,
                           merged['Ad_Spend'] / merged['Total_Revenue'], 0)
merged['T_ROAS'] = np.where(merged['Ad_Spend'] > 0,
                            merged['Total_Revenue'] / merged['Ad_Spend'], 0)
merged['Organic_Ratio'] = np.where(merged['Total_Revenue'] > 0,
                                    merged['Organic_Sales'] / merged['Total_Revenue'], 0)

# ============================================================================
# STEP 7: AGGREGATE BY PLATFORM
# ============================================================================

print("[7/8] Calculating TACoS and T-ROAS by platform...")

def calc_tacos_metrics(subset):
    total_revenue = subset['Total_Revenue'].sum()
    ad_spend = subset['Ad_Spend'].sum()
    ad_sales = subset['Ad_Sales'].sum()
    organic_sales = total_revenue - ad_sales

    return {
        'Total_Revenue': total_revenue,
        'Ad_Spend': ad_spend,
        'Ad_Sales': ad_sales,
        'Organic_Sales': organic_sales,
        'TACoS': (ad_spend / total_revenue * 100) if total_revenue > 0 else 0,
        'T_ROAS': (total_revenue / ad_spend) if ad_spend > 0 else 0,
        'Regular_ROAS': (ad_sales / ad_spend) if ad_spend > 0 else 0,
        'Organic_Ratio': (organic_sales / total_revenue * 100) if total_revenue > 0 else 0,
        'Organic_Lift': (organic_sales / ad_sales * 100) if ad_sales > 0 else 0
    }

perpetua_tacos = calc_tacos_metrics(merged[merged['Advertising_Type'] == 'Perpetua'])
non_perpetua_tacos = calc_tacos_metrics(merged[merged['Advertising_Type'] == 'Non-Perpetua'])

print(f"\n{'='*60}")
print("PERPETUA (with TACoS):")
print(f"{'='*60}")
print(f"  Total Revenue (Orders):     ${perpetua_tacos['Total_Revenue']:,.2f}")
print(f"  Ad-Attributed Sales:        ${perpetua_tacos['Ad_Sales']:,.2f}")
print(f"  Organic Sales:              ${perpetua_tacos['Organic_Sales']:,.2f}")
print(f"  Ad Spend:                   ${perpetua_tacos['Ad_Spend']:,.2f}")
print(f"  ---")
print(f"  Regular ROAS:               {perpetua_tacos['Regular_ROAS']:.2f}x")
print(f"  T-ROAS (Total):             {perpetua_tacos['T_ROAS']:.2f}x")
print(f"  Regular ACOS:               {(perpetua_tacos['Ad_Spend']/perpetua_tacos['Ad_Sales']*100) if perpetua_tacos['Ad_Sales'] > 0 else 0:.1f}%")
print(f"  TACoS (Total):              {perpetua_tacos['TACoS']:.1f}%")
print(f"  Organic Ratio:              {perpetua_tacos['Organic_Ratio']:.1f}%")
print(f"  Organic Lift:               {perpetua_tacos['Organic_Lift']:.0f}%")

print(f"\n{'='*60}")
print("NON-PERPETUA (with TACoS):")
print(f"{'='*60}")
print(f"  Total Revenue (Orders):     ${non_perpetua_tacos['Total_Revenue']:,.2f}")
print(f"  Ad-Attributed Sales:        ${non_perpetua_tacos['Ad_Sales']:,.2f}")
print(f"  Organic Sales:              ${non_perpetua_tacos['Organic_Sales']:,.2f}")
print(f"  Ad Spend:                   ${non_perpetua_tacos['Ad_Spend']:,.2f}")
print(f"  ---")
print(f"  Regular ROAS:               {non_perpetua_tacos['Regular_ROAS']:.2f}x")
print(f"  T-ROAS (Total):             {non_perpetua_tacos['T_ROAS']:.2f}x")
print(f"  Regular ACOS:               {(non_perpetua_tacos['Ad_Spend']/non_perpetua_tacos['Ad_Sales']*100) if non_perpetua_tacos['Ad_Sales'] > 0 else 0:.1f}%")
print(f"  TACoS (Total):              {non_perpetua_tacos['TACoS']:.1f}%")
print(f"  Organic Ratio:              {non_perpetua_tacos['Organic_Ratio']:.1f}%")
print(f"  Organic Lift:               {non_perpetua_tacos['Organic_Lift']:.0f}%")

# ============================================================================
# STEP 8: SAVE PROCESSED DATA
# ============================================================================

print(f"\n[8/8] Saving processed data...")

# Save merged data
merged_file = PROCESSED_DIR / 'orders_advertising_merged.csv'
merged.to_csv(merged_file, index=False)
print(f"  ✓ Saved merged data: {merged_file}")

# Save TACoS summary
import json
tacos_summary = {
    'generated_at': datetime.now().isoformat(),
    'order_files_processed': 2,
    'total_order_lines': len(orders),
    'shipped_orders': len(orders_clean),
    'date_range': f"{min_order_date.date()} to {max_order_date.date()}",
    'perpetua': perpetua_tacos,
    'non_perpetua': non_perpetua_tacos
}

summary_file = OUTPUT_DIR / 'tacos_analysis_summary.json'
with open(summary_file, 'w') as f:
    json.dump(tacos_summary, f, indent=2, default=str)
print(f"  ✓ Saved TACoS summary: {summary_file}")

# Key insights
print(f"\n{'='*100}")
print("KEY INSIGHTS FROM TACoS ANALYSIS")
print(f"{'='*100}")
print()

# Insight 1: Organic ratio comparison
if perpetua_tacos['Organic_Ratio'] > non_perpetua_tacos['Organic_Ratio']:
    print(f"✓ INSIGHT 1: Perpetua has HIGHER organic ratio ({perpetua_tacos['Organic_Ratio']:.1f}% vs {non_perpetua_tacos['Organic_Ratio']:.1f}%)")
    print(f"  Meaning: Perpetua ads are creating MORE organic momentum (flywheel effect)")
else:
    print(f"⚠ INSIGHT 1: Non-Perpetua has higher organic ratio ({non_perpetua_tacos['Organic_Ratio']:.1f}% vs {perpetua_tacos['Organic_Ratio']:.1f}%)")
    print(f"  Meaning: Non-Perpetua products have stronger organic sales base")

print()

# Insight 2: TACoS efficiency
if perpetua_tacos['TACoS'] < non_perpetua_tacos['TACoS']:
    print(f"✓ INSIGHT 2: Perpetua has BETTER TACoS ({perpetua_tacos['TACoS']:.1f}% vs {non_perpetua_tacos['TACoS']:.1f}%)")
    print(f"  Meaning: Perpetua is more efficient when including total business impact")
else:
    print(f"⚠ INSIGHT 2: Non-Perpetua has better TACoS ({non_perpetua_tacos['TACoS']:.1f}% vs {perpetua_tacos['TACoS']:.1f}%)")
    print(f"  Meaning: Non-Perpetua has lower ad dependency")

print()

# Insight 3: Organic lift
print(f"📊 INSIGHT 3: Organic Lift Comparison")
print(f"  Perpetua: {perpetua_tacos['Organic_Lift']:.0f}% organic lift (${perpetua_tacos['Organic_Sales']:,.0f} organic on ${perpetua_tacos['Ad_Sales']:,.0f} ad sales)")
print(f"  Non-Perpetua: {non_perpetua_tacos['Organic_Lift']:.0f}% organic lift (${non_perpetua_tacos['Organic_Sales']:,.0f} organic on ${non_perpetua_tacos['Ad_Sales']:,.0f} ad sales)")

if perpetua_tacos['Organic_Lift'] > non_perpetua_tacos['Organic_Lift']:
    print(f"  ✓ Perpetua drives MORE organic sales per ad dollar (better halo effect)")
else:
    print(f"  ⚠ Non-Perpetua drives more organic sales per ad dollar")

print()

# Insight 4: Total business impact
total_impact_diff = perpetua_tacos['Total_Revenue'] - non_perpetua_tacos['Total_Revenue']
print(f"💰 INSIGHT 4: Total Business Impact")
print(f"  Perpetua total revenue: ${perpetua_tacos['Total_Revenue']:,.0f}")
print(f"  Non-Perpetua total revenue: ${non_perpetua_tacos['Total_Revenue']:,.0f}")
print(f"  Difference: ${total_impact_diff:,.0f} ({'Perpetua generates more' if total_impact_diff > 0 else 'Non-Perpetua generates more'})")

print()
print("=" * 100)
print("✓ ORDER DATA PROCESSING COMPLETE")
print("=" * 100)
print()
print("Next step: Integrate TACoS metrics into Excel dashboard")
