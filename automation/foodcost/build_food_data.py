"""
KFC Food Cost — Data Parser
============================

Reads the source food cost files and produces food_data.json
for the Food Cost Hub dashboard.

Source files expected:
  - FoodCost.xlsx          : Main historical data (Data sheet)
  - CategoriesFC.xlsx      : Ingredient code → category mapping
  - 2026_ΚΟΣΤΟΛΟΓΗΣΗ.xlsx   : Monthly COGS summary (used for cross-validation)
  - ΦΥΡΕΣ_MM_YYYY.xlsx      : Latest month waste analysis (optional)

Run:
    python build_food_data.py [--source-dir DIR] [--output-dir DIR]
"""

import argparse
import json
import sys
from pathlib import Path
from datetime import datetime, timezone

import pandas as pd
from openpyxl import load_workbook


# =====================================================================
# CONFIG
# =====================================================================
KFC_STORES = [
    'KFC ΓΛΥΦΑΔΑ', 'KFC ΠΕΙΡΑΙΑΣ', 'KFC ΡΕΝΤΗΣ', 'KFC ΠΑΓΚΡΑΤΙ', 'KFC ΜΑΡΟΥΣΙ',
    'KFC COSMOS', 'KFC ESCAPE ΙΛΙΟΝ', 'KFC ΚΗΦΙΣΙΑΣ', 'KFC- RIVER WEST',
    'KFC- ΘΩΝ (Αμπελοκηποι)', 'KFC METRO MALL', 'KFC ΣΥΝΤΑΓΜΑ',
    'KFC Smart Park', 'KFC One Salonica', 'KFC Αριστοτέλους',
    'KFC N. ΣΜΥΡΝΗ', 'KFC ΧΑΛΑΝΔΡΙ', 'KFC Ν.ΦΙΛΑΔΕΛΦΕΙΑ',
    'KFC-ΓΕΡΑΚΑΣ', 'KFC-ΜΕΤΑΜΟΡΦΩΣΗ', 'KFC-ΠΑΤΡΑ', 'KFC-ΛΑΡΙΣΑ'
]

# Map FoodCost.xlsx names → canonical short names matching sales dashboard
STORE_NAME_MAP = {
    'KFC ΓΛΥΦΑΔΑ':              'ΓΛΥΦΑΔΑ',
    'KFC ΠΕΙΡΑΙΑΣ':             'ΠΕΙΡΑΙΑΣ',
    'KFC ΡΕΝΤΗΣ':               'ΡΕΝΤΗΣ',
    'KFC ΠΑΓΚΡΑΤΙ':             'ΠΑΓΚΡΑΤΙ',
    'KFC ΜΑΡΟΥΣΙ':              'ΜΑΡΟΥΣΙ',
    'KFC COSMOS':               'COSMOS',
    'KFC ESCAPE ΙΛΙΟΝ':         'ESCAPE',
    'KFC ΚΗΦΙΣΙΑΣ':             'ΚΗΦΙΣΙΑ',
    'KFC- RIVER WEST':          'RIVER',
    'KFC- ΘΩΝ (Αμπελοκηποι)':   'ΘΩΝ',
    'KFC METRO MALL':           'METROMALL',
    'KFC ΣΥΝΤΑΓΜΑ':             'ΣΥΝΤΑΓΜΑ',
    'KFC Smart Park':           'SMART PARK',
    'KFC One Salonica':         'ONE SALONICA',
    'KFC Αριστοτέλους':         'ΑΡΙΣΤΟΤΕΛΟΥΣ',
    'KFC N. ΣΜΥΡΝΗ':            'Ν. ΣΜΥΡΝΗ',
    'KFC ΧΑΛΑΝΔΡΙ':             'ΧΑΛΑΝΔΡΙ',
    'KFC Ν.ΦΙΛΑΔΕΛΦΕΙΑ':        'ΦΙΛΑΔΕΛΦΕΙΑ',
    'KFC-ΓΕΡΑΚΑΣ':              'ΓΕΡΑΚΑΣ',
    'KFC-ΜΕΤΑΜΟΡΦΩΣΗ':          'ΜΕΤΑΜΟΡΦΩΣΗ',
    'KFC-ΠΑΤΡΑ':                'ΠΑΤΡΑ',
    'KFC-ΛΑΡΙΣΑ':               'ΛΑΡΙΣΑ',
}

# Inventory Posting Group classification
# "Pure food" = ingredients that go INTO menu items
# "Total cost" = all of the above + packaging, services, etc.
PURE_FOOD_GROUPS = {'ΑΥΛΕΣ', 'ΠΡΟΙΟΝΤΑ', 'ΕΜΠΟΡΕΥΜΑΤ'}
ALL_FOOD_GROUPS = {'ΑΥΛΕΣ', 'ΠΡΟΙΟΝΤΑ', 'ΕΜΠΟΡΕΥΜΑΤ',
                   'ΥΠΗΡΕΣΙΕΣ', 'ΑΝΑΛΩΣΙΜΑ', 'ΑΝΤΑΛ/ΚΑ', 'SMALLWARES'}


# =====================================================================
# PARSING
# =====================================================================
def load_categories(path):
    """Read CategoriesFC.xlsx → dict {code: category}."""
    print(f"  Loading categories from {path.name}...")
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb['Cat']
    cats = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[2]:
            cats[str(row[0]).strip()] = str(row[2]).strip()
    wb.close()
    print(f"    → {len(cats)} ingredient codes mapped to categories")
    return cats


def load_foodcost_data(path):
    """Read FoodCost.xlsx 'Data' sheet → DataFrame, cleaned and KFC-only."""
    print(f"  Loading historical data from {path.name} ({path.stat().st_size / 1024 / 1024:.1f} MB)...")
    df = pd.read_excel(path, sheet_name='Data', header=2)
    print(f"    → {len(df):,} raw rows loaded")
    
    # Rename columns to short, lowercase keys
    df = df.rename(columns={
        'Έτος': 'year',
        'Μήνας': 'month',
        'Code': 'code',
        'Description': 'desc',
        'Base Unit of Measure': 'unit',
        'Inventory Posting Group': 'group',
        'Reason Code': 'reason',
        'Location Code': 'loc_code',
        'Location Name': 'loc_name',
        'Ideal(Quantity)': 'ideal_qty',
        'Other(Quantity)': 'other_qty',
        'Total(Quantity)': 'total_qty',
        'Ideal(Cost Amount Actual)': 'ideal_cost',
        'Other(Cost Amount Actual)': 'other_cost',
        'Total(Cost Amount Actual)': 'total_cost',
        'Category': 'category_existing',
    })
    
    # Filter to KFC stores only
    before = len(df)
    df = df[df['loc_name'].isin(KFC_STORES)].copy()
    print(f"    → {len(df):,} rows after filtering to {len(KFC_STORES)} KFC stores (dropped {before - len(df):,})")
    
    # Cost values are negative in source (they represent consumption/outflow).
    # Flip sign so positive = consumption €.
    for col in ['ideal_qty','other_qty','total_qty','ideal_cost','other_cost','total_cost']:
        df[col] = -df[col].fillna(0)
    
    # Map store names to canonical short names
    df['store'] = df['loc_name'].map(STORE_NAME_MAP)
    
    return df


def aggregate_monthly(df, categories):
    """
    Build the monthly aggregate structure.
    
    Returns dict {(year, month, store): {pure_food, total_cost, breakdown_by_cat, ...}}
    """
    print("  Aggregating monthly data...")
    
    # Add category from lookup (override existing if any)
    df['category'] = df['code'].map(categories).fillna(df['group'])
    
    # Mark pure-food vs other
    df['is_pure_food'] = df['group'].isin(PURE_FOOD_GROUPS)
    
    # Monthly rollups
    monthly = []
    for (year, month, store), grp in df.groupby(['year', 'month', 'store']):
        pure = grp[grp['is_pure_food']]
        total = grp[grp['group'].isin(ALL_FOOD_GROUPS)]
        
        rec = {
            'year': int(year),
            'month': int(month),
            'store': store,
            
            # Two flavors of food cost (toggleable in dashboard)
            'pure_food': {
                'ideal_cost':  round(float(pure['ideal_cost'].sum()), 2),
                'other_cost':  round(float(pure['other_cost'].sum()), 2),
                'total_cost':  round(float(pure['total_cost'].sum()), 2),
                'ideal_qty':   round(float(pure['ideal_qty'].sum()), 2),
                'total_qty':   round(float(pure['total_qty'].sum()), 2),
            },
            'total': {
                'ideal_cost':  round(float(total['ideal_cost'].sum()), 2),
                'other_cost':  round(float(total['other_cost'].sum()), 2),
                'total_cost':  round(float(total['total_cost'].sum()), 2),
                'ideal_qty':   round(float(total['ideal_qty'].sum()), 2),
                'total_qty':   round(float(total['total_qty'].sum()), 2),
            },
            
            # Waste breakdown (Reason Code = ΦΥΡΑ)
            'waste': {
                'cost': round(float(total[total['reason'] == 'ΦΥΡΑ']['total_cost'].sum()), 2),
                'qty':  round(float(total[total['reason'] == 'ΦΥΡΑ']['total_qty'].sum()), 2),
            },
            
            # Reason breakdown (cost by reason, food relevant)
            'by_reason': {
                str(r) if r is not None and pd.notna(r) else 'NORMAL':
                    round(float(total[total['reason'].fillna('NORMAL').replace({None: 'NORMAL'}) == 
                                       (r if pd.notna(r) else 'NORMAL')]['total_cost'].sum()), 2)
                for r in total['reason'].fillna('NORMAL').unique()
            },
            
            # Category breakdown (pure food)
            'by_category': pure.groupby('category').agg(
                ideal_cost=('ideal_cost', 'sum'),
                other_cost=('other_cost', 'sum'),
                total_cost=('total_cost', 'sum'),
            ).round(2).to_dict(orient='index'),
        }
        monthly.append(rec)
    
    print(f"    → {len(monthly)} (year × month × store) records")
    return monthly


def aggregate_top_variances(df, categories, year, month):
    """
    For a given month, compute top-variance ingredients per store.
    Returns dict {store: [{code, desc, category, ideal_cost, other_cost, total_cost}, ...]}
    """
    df['category'] = df['code'].map(categories).fillna(df['group'])
    pure = df[df['group'].isin(PURE_FOOD_GROUPS)].copy()
    
    target = pure[(pure['year'] == year) & (pure['month'] == month)].copy()
    
    result = {}
    for store, grp in target.groupby('store'):
        items = grp.groupby(['code','desc','category']).agg(
            ideal_cost=('ideal_cost', 'sum'),
            other_cost=('other_cost', 'sum'),
            total_cost=('total_cost', 'sum'),
            total_qty=('total_qty', 'sum'),
        ).reset_index()
        items['abs_other'] = items['other_cost'].abs()
        # Top 20 by absolute variance
        top = items.nlargest(20, 'abs_other')
        result[store] = [
            {
                'code': str(row['code']),
                'desc': str(row['desc']),
                'category': str(row['category']),
                'ideal_cost': round(float(row['ideal_cost']), 2),
                'other_cost': round(float(row['other_cost']), 2),
                'total_cost': round(float(row['total_cost']), 2),
                'total_qty': round(float(row['total_qty']), 2),
            }
            for _, row in top.iterrows()
        ]
    
    return result


def load_kostologisi(path):
    """
    Read 2026_ΚΟΣΤΟΛΟΓΗΣΗ.xlsx → for cross-validation.
    Returns dict {(year, month, store): {αρχικο, αγορες, τελικο, COGS}}
    """
    print(f"  Loading ΚΟΣΤΟΛΟΓΗΣΗ summary from {path.name}...")
    wb = load_workbook(path, data_only=True)
    ws = wb['ΚΟΣΤΟΛΟΓΗΣΗ']
    
    # Detect month column starts (row 2 has period headers like "Food Cost _ JAN 2026")
    period_starts = []
    for cell in ws[2]:
        if cell.value and 'Food Cost' in str(cell.value) and 'ALL' not in str(cell.value):
            period_starts.append((cell.column, str(cell.value)))
    
    # Parse month from "Food Cost _JAN 2026" or "Food Cost _ FEB 2026"
    MONTHS = {'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,
              'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
    
    result = {}
    for col_start, header in period_starts:
        # Extract month + year from header
        parts = header.replace('Food Cost', '').replace('_', '').strip().split()
        month_str = parts[0].strip()
        year = int(parts[-1])
        if month_str not in MONTHS:
            continue
        month = MONTHS[month_str]
        
        # Read 22 store rows (4-25)
        for r in range(4, 26):
            store_name = ws.cell(row=r, column=col_start).value  # "[201] KFC-ΓΛΥΦΑΔΑ"
            cogs = ws.cell(row=r, column=col_start + 5).value  # Κόστος Πωληθέντων
            if store_name and cogs is not None:
                result[(year, month, str(store_name).strip())] = {
                    'cogs': float(cogs)
                }
    wb.close()
    print(f"    → {len(result)} period × store records loaded")
    return result


# =====================================================================
# BUILD JSON
# =====================================================================
def build_food_data_json(source_dir: Path, output_dir: Path):
    """Main pipeline: read sources, aggregate, write food_data.json."""
    print("=" * 60)
    print("KFC Food Cost Hub — Data Build")
    print("=" * 60)
    print(f"Source: {source_dir}")
    print(f"Output: {output_dir}")
    print()
    
    # Step 1: Load lookups
    print("STEP 1: Load reference data")
    categories = load_categories(source_dir / 'CategoriesFC.xlsx')
    
    # Step 2: Load main data
    print("\nSTEP 2: Load main historical data")
    df = load_foodcost_data(source_dir / 'FoodCost.xlsx')
    
    # Step 3: Optional cross-validation
    print("\nSTEP 3: Load ΚΟΣΤΟΛΟΓΗΣΗ for cross-check (optional)")
    kostologisi = {}
    kosto_files = list(source_dir.glob('*ΚΟΣΤΟΛΟΓΗΣΗ*.xlsx'))
    if kosto_files:
        kostologisi = load_kostologisi(kosto_files[0])
    else:
        print("    → No ΚΟΣΤΟΛΟΓΗΣΗ file found, skipping")
    
    # Step 4: Aggregate
    print("\nSTEP 4: Aggregate monthly")
    monthly_records = aggregate_monthly(df, categories)
    
    # Step 5: Top variance ingredients for latest month
    print("\nSTEP 5: Compute top-variance ingredients (latest month)")
    latest_year = int(df['year'].max())
    latest_month = int(df[df['year'] == latest_year]['month'].max())
    print(f"    Latest period: {latest_year}-{latest_month:02d}")
    top_variances = aggregate_top_variances(df, categories, latest_year, latest_month)
    
    # Step 6: Build final JSON
    print("\nSTEP 6: Build JSON output")
    
    # Collect available periods
    periods = sorted({(r['year'], r['month']) for r in monthly_records})
    period_strings = [f"{y}-{m:02d}" for y, m in periods]
    
    # Build categories list with totals
    all_categories = sorted(set(categories.values()))
    
    output = {
        'meta': {
            'generated_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
            'source_files': [
                'FoodCost.xlsx',
                'CategoriesFC.xlsx',
                next((f.name for f in kosto_files), None),
            ],
            'kfc_stores': [STORE_NAME_MAP[s] for s in KFC_STORES],
            'periods': period_strings,
            'latest_period': f"{latest_year}-{latest_month:02d}",
            'categories': all_categories,
        },
        'monthly': monthly_records,
        'top_variances_latest': top_variances,
        'kostologisi': [
            {'year': y, 'month': m, 'store_raw': s, 'cogs': v['cogs']}
            for (y, m, s), v in kostologisi.items()
        ] if kostologisi else [],
    }
    
    # Write
    output_path = output_dir / 'food_data.json'
    output_dir.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=None, separators=(',',':'))
    
    size_kb = output_path.stat().st_size / 1024
    print(f"\n✓ Wrote {output_path.name} ({size_kb:.1f} KB)")
    print(f"  - {len(monthly_records)} monthly records")
    print(f"  - {len(periods)} periods covered: {period_strings[0]} → {period_strings[-1]}")
    print(f"  - {len(all_categories)} categories")
    print(f"  - {len(top_variances)} stores with variance breakdown for {latest_year}-{latest_month:02d}")
    
    return output


# =====================================================================
# CLI
# =====================================================================
def main():
    parser = argparse.ArgumentParser(description='Build food_data.json for the Food Cost Hub')
    parser.add_argument('--source-dir', default='/mnt/user-data/uploads',
                        help='Folder containing source xlsx files')
    parser.add_argument('--output-dir', default='/mnt/user-data/outputs/food-cost-hub',
                        help='Output folder for food_data.json')
    args = parser.parse_args()
    
    source = Path(args.source_dir)
    output = Path(args.output_dir)
    
    if not source.exists():
        print(f"ERROR: Source directory not found: {source}", file=sys.stderr)
        sys.exit(1)
    
    build_food_data_json(source, output)


if __name__ == '__main__':
    main()
