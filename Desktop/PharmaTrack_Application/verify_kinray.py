import pandas as pd
import sys

path = sys.argv[1]  # pass kinray csv path
df = pd.read_csv(path, dtype=str)

print("=== KINRAY VERIFICATION ===")
print(f"Total rows: {len(df)}")
print(f"Columns: {df.columns.tolist()}")

# Find price column
price_col = None
for col in df.columns:
    if any(x in col.lower() for x in
           ['invoice', 'price', 'amount']):
        price_col = col
        print(f"Price column found: {col}")
        break

if price_col:
    df['__price__'] = pd.to_numeric(
        df[price_col].astype(str)
        .str.replace(',', '', regex=False)
        .str.replace('$', '', regex=False)
        .str.replace(r'[^0-9.\-]', '', regex=True),
        errors='coerce'
    ).fillna(0)
    print(f"Total purchased: ${df['__price__'].sum():,.2f}")

# NDC column
ndc_col = next((c for c in df.columns if 'ndc' in c.lower() or 'upc' in c.lower()), None)
if ndc_col:
    print(f"Unique NDCs: {df[ndc_col].nunique()}")

# Date column
date_col = next((c for c in df.columns if 'date' in c.lower()), None)
if date_col:
    print(f"Date range: {df[date_col].min()} to {df[date_col].max()}")
