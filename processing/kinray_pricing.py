# Kinray unit price lookups: find_kinray_price_by_month() — month-aware price search with forward/backward fallback.

import pandas as pd


def find_kinray_price_by_month(ndc, fill_date, kinray_df):
    """
    Find Kinray price for NDC based on fill_date:
    1. Search same month/year as fill_date (latest purchase in that month)
    2. If not found, search backwards month by month
    3. If not found, search forwards month by month
    4. Return 0 if never found
    """
    if pd.isna(fill_date) or kinray_df.empty:
        return 0

    # Filter for this NDC
    ndc_purchases = kinray_df[kinray_df['NDC #'] == ndc].copy()
    if ndc_purchases.empty:
        return 0

    # Ensure DATE is datetime
    ndc_purchases['DATE'] = pd.to_datetime(ndc_purchases['DATE'], errors='coerce')
    ndc_purchases = ndc_purchases.dropna(subset=['DATE', '__UnitPrice__'])

    if ndc_purchases.empty:
        return 0

    fill_date = pd.to_datetime(fill_date)
    target_year = fill_date.year
    target_month = fill_date.month

    # Try same month first
    same_month = ndc_purchases[
        (ndc_purchases['DATE'].dt.year == target_year) &
        (ndc_purchases['DATE'].dt.month == target_month)
    ]
    if not same_month.empty:
        return same_month.sort_values('DATE').iloc[-1]['__UnitPrice__']

    # Get min and max dates available
    min_date = ndc_purchases['DATE'].min()
    max_date = ndc_purchases['DATE'].max()

    # Search backwards
    current_date = fill_date
    while current_date >= min_date:
        current_date = current_date - pd.DateOffset(months=1)
        month_data = ndc_purchases[
            (ndc_purchases['DATE'].dt.year == current_date.year) &
            (ndc_purchases['DATE'].dt.month == current_date.month)
        ]
        if not month_data.empty:
            return month_data.sort_values('DATE').iloc[-1]['__UnitPrice__']

    # Search forwards
    current_date = fill_date
    while current_date <= max_date:
        current_date = current_date + pd.DateOffset(months=1)
        month_data = ndc_purchases[
            (ndc_purchases['DATE'].dt.year == current_date.year) &
            (ndc_purchases['DATE'].dt.month == current_date.month)
        ]
        if not month_data.empty:
            return month_data.sort_values('DATE').iloc[-1]['__UnitPrice__']

    return 0
