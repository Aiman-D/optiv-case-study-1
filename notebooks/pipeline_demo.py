# CS1 Notebook: Vendor AI Usage Analysis Pipeline
# Run cells sequentially to execute the full pipeline.

# ── Cell 1: Setup ──
import sys, os
sys.path.insert(0, '../src')
os.chdir('..')

# ── Cell 2: M1 — Criticality Engine ──
from criticality import evaluate_vendor, load_rubric
import pandas as pd

rubric = load_rubric('config/rubric.yaml')
df = pd.read_excel('data/Meridian_Vendor_Input.xlsx', sheet_name='Vendor Inventory', header=3)
df = df[df['Vendor ID'].str.startswith('V-', na=False) & (df['Vendor ID'] != 'V-000')]

print("=" * 70)
print("M1: CRITICALITY ENGINE RESULTS")
print("=" * 70)
for idx, row in df.iterrows():
    tier, rationale, depth = evaluate_vendor(row, rubric)
    print(f"\n{row['Vendor ID']} — {row['Vendor Name']}")
    print(f"  Tier: {tier}")
    print(f"  Depth: {depth}")

# ── Cell 3: M2 — OSINT Collectors ──
# NOTE: This crawls live vendor websites. Skip if you already have cached data.
# from collectors import OSINTCollector
# collector = OSINTCollector('config/sources.yaml')
# ... (see src/collectors.py)

# ── Cell 4: M3 — Classifier ──
from classify import classify_vendor
from pathlib import Path

print("\n" + "=" * 70)
print("M3: CLASSIFIER RESULTS")
print("=" * 70)
cache_dir = Path('cache')
for vendor_dir in sorted(cache_dir.iterdir()):
    if vendor_dir.is_dir() and vendor_dir.name.startswith('V-'):
        signals = classify_vendor(vendor_dir.name)

# ── Cell 5: M4+M5 — Risk Engine & Export ──
from risk_export import process_and_export
print("\n" + "=" * 70)
print("M4+M5: RISK ENGINE & EXPORT")
print("=" * 70)
process_and_export()

# ── Cell 6: View Results ──
df_out = pd.read_excel('data/Meridian_Vendor_Output.xlsx', sheet_name='Vendor Inventory', header=3)
cols = ['Vendor ID', 'Vendor Name', 'Criticality Tier', 'AI Usage Detected (Y/N)', 'AI Security Risk Classification']
print("\n" + "=" * 70)
print("FINAL OUTPUT SUMMARY")
print("=" * 70)
print(df_out[cols].to_string(index=False))
