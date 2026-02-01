#!/usr/bin/env python3
"""
Refresh All Reports - Automation Script
Re-runs entire analysis pipeline with updated data files
"""

import subprocess
import sys
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent
SCRIPTS_DIR = BASE_DIR / 'scripts'

print("=" * 80)
print("PERPETUA REPORT REFRESH - AUTOMATED PIPELINE")
print("=" * 80)
print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print()

scripts_to_run = [
    ('1_process_campaign_data.py', 'Processing campaign data and tagging Perpetua vs Non-Perpetua'),
    ('2_asin_level_analysis.py', 'Running ASIN-level performance analysis'),
    ('3_generate_performance_report.py', 'Generating performance reports and visualizations'),
    ('4_generate_excel_dashboard.py', 'Creating Excel dashboard')
]

failed_scripts = []

for idx, (script, description) in enumerate(scripts_to_run, 1):
    print(f"[{idx}/{len(scripts_to_run)}] {description}...")
    print(f"  Running: {script}")

    script_path = SCRIPTS_DIR / script

    if not script_path.exists():
        print(f"  ✗ ERROR: Script not found: {script_path}")
        failed_scripts.append(script)
        continue

    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=300  # 5 minutes max per script
        )

        if result.returncode == 0:
            print(f"  ✓ {script} completed successfully")
        else:
            print(f"  ✗ {script} failed with exit code {result.returncode}")
            print(f"  Error output:")
            print(result.stderr)
            failed_scripts.append(script)

    except subprocess.TimeoutExpired:
        print(f"  ✗ {script} timed out after 5 minutes")
        failed_scripts.append(script)
    except Exception as e:
        print(f"  ✗ {script} failed with error: {e}")
        failed_scripts.append(script)

    print()

print("=" * 80)
print("REFRESH COMPLETE")
print("=" * 80)
print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print()

if failed_scripts:
    print(f"⚠ WARNING: {len(failed_scripts)} script(s) failed:")
    for script in failed_scripts:
        print(f"  - {script}")
    print()
    print("Review errors above and re-run failed scripts manually")
    sys.exit(1)
else:
    print("✓ All reports refreshed successfully!")
    print()
    print("Updated files in outputs/:")
    print("  - Perpetua_Performance_Dashboard_YYYYMMDD.xlsx")
    print("  - Campaign_Performance_Report.txt")
    print("  - Campaign_Performance_Summary.md")
    print("  - 4 visualization PNGs")
    print()
    print("To refresh again in the future, run:")
    print("  python scripts/refresh_reports.py")
