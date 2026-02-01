#!/usr/bin/env python3
"""
Generate Campaign Performance Analysis Report with Visualizations
"""

import pandas as pd
import json
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from pathlib import Path
from datetime import datetime

# Paths
BASE_DIR = Path(__file__).parent.parent
AGG_DIR = BASE_DIR / 'data' / 'aggregated'
PROCESSED_DIR = BASE_DIR / 'data' / 'processed'
OUTPUT_DIR = BASE_DIR / 'outputs'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print("=" * 80)
print("GENERATING CAMPAIGN PERFORMANCE ANALYSIS REPORT")
print("=" * 80)
print()

# Load comparison data
print("[1/4] Loading analysis results...")
with open(AGG_DIR / 'asin_level_comparison.json', 'r') as f:
    analysis = json.load(f)

comparison_df = pd.read_csv(AGG_DIR / 'asin_comparison_full.csv', index_col=0)
print("  ✓ Data loaded")
print()

# Extract metrics
perpetua = analysis['perpetua_metrics']
non_perpetua = analysis['non_perpetua_metrics']

#  Generate visualizations
print("[2/4] Creating visualizations...")

# Set style
plt.style.use('default')
colors = {'Perpetua': '#2E86AB', 'Non-Perpetua': '#A23B72'}

# Figure 1: Perpetua vs Non-Perpetua Comparison (Side-by-side bars)
fig, axes = plt.subplots(2, 3, figsize=(15, 10))
fig.suptitle('Perpetua vs Non-Perpetua: Performance Comparison', fontsize=16, fontweight='bold')

metrics_to_plot = [
    ('ACOS', 'ACOS (lower is better)', True),  # True = lower is better
    ('ROAS', 'ROAS (higher is better)', False),
    ('Avg_CPC', 'Avg CPC ($)', True),
    ('Avg_CVR', 'Conversion Rate', False),
    ('CTR', 'Click-Through Rate', False),
    ('Total_Spend', 'Total Spend ($)', None)  # None = neutral
]

for idx, (metric, label, lower_better) in enumerate(metrics_to_plot):
    ax = axes[idx // 3, idx % 3]

    values = [perpetua[metric], non_perpetua[metric]]
    labels_bar = ['Perpetua', 'Non-Perpetua']

    bars = ax.bar(labels_bar, values, color=[colors['Perpetua'], colors['Non-Perpetua']])
    ax.set_ylabel(label)
    ax.set_title(label)
    ax.grid(axis='y', alpha=0.3)

    # Add value labels on bars
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.2f}' if height < 1000 else f'${height:,.0f}',
                ha='center', va='bottom')

    # Highlight better performer
    if lower_better is not None:
        better_idx = 0 if (perpetua[metric] < non_perpetua[metric] and lower_better) or \
                          (perpetua[metric] > non_perpetua[metric] and not lower_better) else 1
        bars[better_idx].set_edgecolor('green')
        bars[better_idx].set_linewidth(3)

plt.tight_layout()
plt.savefig(OUTPUT_DIR / 'perpetua_vs_nonperpetua_comparison.png', dpi=300, bbox_inches='tight')
print("  ✓ Saved: perpetua_vs_nonperpetua_comparison.png")

# Figure 2: Spend vs Sales scatter plot
fig, ax = plt.subplots(figsize=(10, 6))
ax.scatter([perpetua['Total_Spend']], [perpetua['Total_Sales']],
           s=500, c=colors['Perpetua'], alpha=0.6, label='Perpetua', edgecolors='black')
ax.scatter([non_perpetua['Total_Spend']], [non_perpetua['Total_Sales']],
           s=500, c=colors['Non-Perpetua'], alpha=0.6, label='Non-Perpetua', edgecolors='black')

ax.set_xlabel('Total Spend ($)', fontsize=12)
ax.set_ylabel('Total Sales ($)', fontsize=12)
ax.set_title('Spend vs Sales: Perpetua vs Non-Perpetua', fontsize=14, fontweight='bold')
ax.legend()
ax.grid(alpha=0.3)

# Add diagonal line for ROAS = 1
max_val = max(perpetua['Total_Spend'], non_perpetua['Total_Spend'],
              perpetua['Total_Sales'], non_perpetua['Total_Sales'])
ax.plot([0, max_val], [0, max_val], 'k--', alpha=0.3, label='Break-even (ROAS=1)')

plt.tight_layout()
plt.savefig(OUTPUT_DIR / 'spend_vs_sales_scatter.png', dpi=300, bbox_inches='tight')
print("  ✓ Saved: spend_vs_sales_scatter.png")

# Figure 3: ROAS Comparison with target line
fig, ax = plt.subplots(figsize=(10, 6))
x_pos = [0, 1]
roas_values = [perpetua['ROAS'], non_perpetua['ROAS']]
bars = ax.bar(x_pos, roas_values, color=[colors['Perpetua'], colors['Non-Perpetua']],
              edgecolor='black', linewidth=1.5)

# Add target line at ROAS = 2.0
ax.axhline(y=2.0, color='green', linestyle='--', linewidth=2, label='Target ROAS (2.0)')

ax.set_xticks(x_pos)
ax.set_xticklabels(['Perpetua', 'Non-Perpetua'])
ax.set_ylabel('ROAS', fontsize=12)
ax.set_title('Return on Ad Spend (ROAS) Comparison', fontsize=14, fontweight='bold')
ax.legend()
ax.grid(axis='y', alpha=0.3)

# Add value labels
for i, (bar, val) in enumerate(zip(bars, roas_values)):
    ax.text(bar.get_x() + bar.get_width()/2., val + 0.05,
            f'{val:.2f}', ha='center', va='bottom', fontsize=12, fontweight='bold')

    # Add performance indicator
    if val >= 2.0:
        symbol = "✓ Meets Target"
        color_text = 'green'
    else:
        symbol = "✗ Below Target"
        color_text = 'red'
    ax.text(bar.get_x() + bar.get_width()/2., 0.1,
            symbol, ha='center', va='bottom', fontsize=10, color=color_text)

plt.tight_layout()
plt.savefig(OUTPUT_DIR / 'roas_comparison.png', dpi=300, bbox_inches='tight')
print("  ✓ Saved: roas_comparison.png")

# Figure 4: Efficiency metrics (ACOS, CPC, Conversion Rate)
fig, axes = plt.subplots(1, 3, figsize=(15, 5))
fig.suptitle('Efficiency Metrics Comparison', fontsize=16, fontweight='bold')

efficiency_metrics = [
    ('ACOS', 'ACOS (%)', 100),  # multiply by 100 for percentage
    ('Avg_CPC', 'Cost Per Click ($)', 1),
    ('Avg_CVR', 'Conversion Rate (%)', 100)
]

for idx, (metric, ylabel, multiplier) in enumerate(efficiency_metrics):
    ax = axes[idx]
    values = [perpetua[metric] * multiplier, non_perpetua[metric] * multiplier]
    bars = ax.bar(['Perpetua', 'Non-Perpetua'], values,
                   color=[colors['Perpetua'], colors['Non-Perpetua']])
    ax.set_ylabel(ylabel)
    ax.set_title(ylabel)
    ax.grid(axis='y', alpha=0.3)

    # Highlight better performer (lower is better for all)
    better_idx = 0 if values[0] < values[1] else 1
    bars[better_idx].set_edgecolor('green')
    bars[better_idx].set_linewidth(3)

    # Add value labels
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.2f}', ha='center', va='bottom', fontweight='bold')

plt.tight_layout()
plt.savefig(OUTPUT_DIR / 'efficiency_metrics_comparison.png', dpi=300, bbox_inches='tight')
print("  ✓ Saved: efficiency_metrics_comparison.png")

print()

# Generate text report
print("[3/4] Generating text report...")

report_lines = []
report_lines.append("=" * 80)
report_lines.append("CAMPAIGN PERFORMANCE ANALYSIS: PERPETUA VS NON-PERPETUA")
report_lines.append("=" * 80)
report_lines.append("")
report_lines.append(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
report_lines.append("")

report_lines.append("EXECUTIVE SUMMARY")
report_lines.append("-" * 80)
report_lines.append(f"  Perpetua ASINs Analyzed: {analysis['perpetua_asins_analyzed']}")
report_lines.append(f"  Non-Perpetua ASINs Analyzed: {analysis['non_perpetua_asins_analyzed']}")
report_lines.append("")

report_lines.append("PERFORMANCE METRICS")
report_lines.append("-" * 80)
report_lines.append(f"{'Metric':<25} | {'Perpetua':>15} | {'Non-Perpetua':>15} | {'Winner':<12}")
report_lines.append("-" * 80)

metrics_display = [
    ('Total_Spend', 'Total Spend', '$', False),
    ('Total_Sales', 'Total Sales', '$', False),
    ('Total_Orders', 'Total Orders', '', False),
    ('ACOS', 'ACOS', '%', True),  # True = lower is better
    ('ROAS', 'ROAS', 'x', False),
    ('Avg_CPC', 'Avg CPC', '$', True),
    ('Avg_CVR', 'Conversion Rate', '%', False),
    ('CTR', 'CTR', '%', False)
]

for metric_key, metric_name, unit, lower_better in metrics_display:
    p_val = perpetua[metric_key]
    np_val = non_perpetua[metric_key]

    # Format values
    if unit == '$':
        p_str = f"${p_val:,.2f}"
        np_str = f"${np_val:,.2f}"
    elif unit == '%':
        p_str = f"{p_val*100:.2f}%"
        np_str = f"{np_val*100:.2f}%"
    elif unit == 'x':
        p_str = f"{p_val:.2f}x"
        np_str = f"{np_val:.2f}x"
    else:
        p_str = f"{p_val:,.0f}"
        np_str = f"{np_val:,.0f}"

    # Determine winner
    if lower_better:
        winner = "Perpetua ✓" if p_val < np_val else "Non-Perpetua ✓"
    else:
        winner = "Perpetua ✓" if p_val > np_val else "Non-Perpetua ✓"

    report_lines.append(f"{metric_name:<25} | {p_str:>15} | {np_str:>15} | {winner:<12}")

report_lines.append("")
report_lines.append("KEY INSIGHTS")
report_lines.append("-" * 80)

# Calculate improvement percentages
acos_improvement = ((non_perpetua['ACOS'] - perpetua['ACOS']) / non_perpetua['ACOS'] * 100) if non_perpetua['ACOS'] > 0 else 0
roas_improvement = ((perpetua['ROAS'] - non_perpetua['ROAS']) / non_perpetua['ROAS'] * 100) if non_perpetua['ROAS'] > 0 else 0
cpc_improvement = ((non_perpetua['Avg_CPC'] - perpetua['Avg_CPC']) / non_perpetua['Avg_CPC'] * 100) if non_perpetua['Avg_CPC'] > 0 else 0
cvr_improvement = ((perpetua['Avg_CVR'] - non_perpetua['Avg_CVR']) / non_perpetua['Avg_CVR'] * 100) if non_perpetua['Avg_CVR'] > 0 else 0

insights = []

if acos_improvement > 0:
    insights.append(f"  ✓ Perpetua has {acos_improvement:.1f}% better ACOS (more efficient)")
else:
    insights.append(f"  ⚠ Non-Perpetua has {abs(acos_improvement):.1f}% better ACOS")

if roas_improvement > 0:
    insights.append(f"  ✓ Perpetua has {roas_improvement:.1f}% better ROAS")
else:
    insights.append(f"  ⚠ Non-Perpetua has {abs(roas_improvement):.1f}% better ROAS")

if cpc_improvement > 0:
    insights.append(f"  ✓ Perpetua has {cpc_improvement:.1f}% lower CPC (more cost-efficient)")
else:
    insights.append(f"  ⚠ Non-Perpetua has {abs(cpc_improvement):.1f}% lower CPC")

# Spend analysis
spend_diff_pct = ((perpetua['Total_Spend'] - non_perpetua['Total_Spend']) / non_perpetua['Total_Spend'] * 100)
insights.append(f"  • Perpetua spends {spend_diff_pct:.0f}% more than non-Perpetua products")

# Sales analysis
sales_diff_pct = ((perpetua['Total_Sales'] - non_perpetua['Total_Sales']) / non_perpetua['Total_Sales'] * 100)
insights.append(f"  • Perpetua generates {sales_diff_pct:.0f}% more sales revenue")

# ASIN analysis
insights.append(f"  • Perpetua manages {analysis['perpetua_asins_analyzed']} ASINs vs {analysis['non_perpetua_asins_analyzed']} non-Perpetua ASINs")

# Per-ASIN efficiency
spend_per_asin_p = perpetua['Total_Spend'] / analysis['perpetua_asins_analyzed']
spend_per_asin_np = non_perpetua['Total_Spend'] / analysis['non_perpetua_asins_analyzed']
sales_per_asin_p = perpetua['Total_Sales'] / analysis['perpetua_asins_analyzed']
sales_per_asin_np = non_perpetua['Total_Sales'] / analysis['non_perpetua_asins_analyzed']

insights.append("")
insights.append("Per-ASIN Metrics:")
insights.append(f"  • Perpetua: ${spend_per_asin_p:,.0f} spend/ASIN → ${sales_per_asin_p:,.0f} sales/ASIN")
insights.append(f"  • Non-Perpetua: ${spend_per_asin_np:,.0f} spend/ASIN → ${sales_per_asin_np:,.0f} sales/ASIN")

report_lines.extend(insights)

report_lines.append("")
report_lines.append("RECOMMENDATIONS")
report_lines.append("-" * 80)

recommendations = []
if roas_improvement < 0:
    recommendations.append("  1. Non-Perpetua products show better ROAS - investigate what strategies can be")
    recommendations.append("     applied from non-Perpetua to Perpetua campaigns")

if perpetua['ROAS'] < 2.0:
    recommendations.append("  2. Perpetua ROAS is below 2.0 target - consider:")
    recommendations.append("     - Streams bid optimization adjustments")
    recommendations.append("     - Negative keyword expansion to reduce wasted spend")
    recommendations.append("     - Budget reallocation from low-performing to high-performing campaigns")

if perpetua['Avg_CPC'] > non_perpetua['Avg_CPC']:
    recommendations.append("  3. Perpetua has higher CPC - this may indicate:")
    recommendations.append("     - More competitive keywords (expected for higher-volume products)")
    recommendations.append("     - Opportunity to refine bidding strategies")

recommendations.append("  4. Continue leveraging automation for Perpetua products while monitoring efficiency")
recommendations.append("  5. Consider expanding Perpetua management to high-performing non-Perpetua ASINs")

report_lines.extend(recommendations)

report_lines.append("")
report_lines.append("=" * 80)
report_lines.append("END OF REPORT")
report_lines.append("=" * 80)

# Save text report
report_file = OUTPUT_DIR / 'Campaign_Performance_Report.txt'
with open(report_file, 'w') as f:
    f.write('\n'.join(report_lines))

print(f"  ✓ Saved: {report_file.name}")
print()

# Save summary for markdown
print("[4/4] Creating markdown summary...")

md_lines = []
md_lines.append("# Campaign Performance Analysis: Perpetua vs Non-Perpetua")
md_lines.append("")
md_lines.append(f"**Report Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
md_lines.append("")

md_lines.append("## Executive Summary")
md_lines.append("")
md_lines.append(f"- **Perpetua ASINs Analyzed:** {analysis['perpetua_asins_analyzed']}")
md_lines.append(f"- **Non-Perpetua ASINs Analyzed:** {analysis['non_perpetua_asins_analyzed']}")
md_lines.append("")

md_lines.append("## Performance Comparison")
md_lines.append("")
md_lines.append("| Metric | Perpetua | Non-Perpetua | Winner |")
md_lines.append("|--------|----------|--------------|--------|")

for metric_key, metric_name, unit, lower_better in metrics_display:
    p_val = perpetua[metric_key]
    np_val = non_perpetua[metric_key]

    if unit == '$':
        p_str = f"\\${p_val:,.2f}"
        np_str = f"\\${np_val:,.2f}"
    elif unit == '%':
        p_str = f"{p_val*100:.2f}%"
        np_str = f"{np_val*100:.2f}%"
    elif unit == 'x':
        p_str = f"{p_val:.2f}x"
        np_str = f"{np_val:.2f}x"
    else:
        p_str = f"{p_val:,.0f}"
        np_str = f"{np_val:,.0f}"

    if lower_better:
        winner = "Perpetua ✓" if p_val < np_val else "Non-Perpetua ✓"
    else:
        winner = "Perpetua ✓" if p_val > np_val else "Non-Perpetua ✓"

    md_lines.append(f"| {metric_name} | {p_str} | {np_str} | {winner} |")

md_lines.append("")
md_lines.append("## Key Insights")
md_lines.append("")
for insight in insights:
    md_lines.append(insight)

md_lines.append("")
md_lines.append("## Visualizations")
md_lines.append("")
md_lines.append("### Perpetua vs Non-Perpetua Comparison")
md_lines.append("![Comparison](perpetua_vs_nonperpetua_comparison.png)")
md_lines.append("")
md_lines.append("### Spend vs Sales")
md_lines.append("![Spend vs Sales](spend_vs_sales_scatter.png)")
md_lines.append("")
md_lines.append("### ROAS Comparison")
md_lines.append("![ROAS](roas_comparison.png)")
md_lines.append("")
md_lines.append("### Efficiency Metrics")
md_lines.append("![Efficiency](efficiency_metrics_comparison.png)")

md_file = OUTPUT_DIR / 'Campaign_Performance_Summary.md'
with open(md_file, 'w') as f:
    f.write('\n'.join(md_lines))

print(f"  ✓ Saved: {md_file.name}")
print()

print("=" * 80)
print("✓ PERFORMANCE REPORT GENERATION COMPLETE")
print("=" * 80)
print()
print(f"Generated files in {OUTPUT_DIR}:")
print("  - perpetua_vs_nonperpetua_comparison.png")
print("  - spend_vs_sales_scatter.png")
print("  - roas_comparison.png")
print("  - efficiency_metrics_comparison.png")
print("  - Campaign_Performance_Report.txt")
print("  - Campaign_Performance_Summary.md")
