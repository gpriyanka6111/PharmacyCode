# Vendor file (Kinray + extras) reading, NDC normalization, and ship-qty / price pivot aggregation.

import os
import re

import numpy as np
import pandas as pd


def parse_vendor_files(vendor_paths):
    """
    Read every vendor CSV in vendor_paths, normalise headers/NDC/Shipped/PRICE,
    stash Kinray rows for month-aware price lookups, and build the qty + price pivots.

    Returns
    -------
    vendor_pivot        : DataFrame  – shipped qty  by NDC × Vendor (columns uppercased)
    vendor_price_pivot  : DataFrame  – last PRICE   by NDC × Vendor (columns as VENDOR_PRICE)
    vendor_names        : list[str]  – uppercased vendor labels (matches vendor_pivot cols)
    kinray_all          : DataFrame  – every Kinray row (NDC #, DATE, OURCASE, Shipped, PRICE, __UnitPrice__)
    kinray_latest       : DataFrame  – one row per NDC with Kinray_UPrice
    all_vendor_df       : DataFrame  – combined rows (NDC #, Shipped, PRICE, DATE, Vendor, Drug Name)
    """
    all_vendor_rows = []   # for Alternate NDC logic

    vendor_frames_qty, vendor_frames_price, vendor_names = [], [], []
    kinray_rows = []   # stash rows (NDC, DATE, UnitPrice) for Kinray only

    def _norm_headers(cols):
        # collapse whitespace and convert NBSP to space
        return [re.sub(r'\s+', ' ', str(c).replace('\xa0', ' ')).strip() for c in cols]

    def _pick(lower_to_orig, cands):
        # exact case-insensitive first
        for c in cands:
            if c in lower_to_orig:
                return lower_to_orig[c]
        # partial contains fallback
        for c in cands:
            for low, orig in lower_to_orig.items():
                if c in low:
                    return orig
        return None

    for i, vp in enumerate(vendor_paths, start=1):
        raw = pd.read_csv(vp, dtype=str)

        # Normalize headers
        raw.columns = _norm_headers(raw.columns)
        lower_to_orig = {c.lower(): c for c in raw.columns}

        # Required columns (robust to vendor variations)
        ndc_col = _pick(lower_to_orig, ['ndc/upc', 'ndc #', 'ndc', 'ndc#', 'ndc number',
                        'ndc no', 'ndc upc', 'ndcupc', 'NDC/UPC', 'Item NDC/UPC (Current)'])
        ship_col = _pick(lower_to_orig, ['ship qty', 'shipped', 'shipped qty', 'qty shipped',
                         'quantity shipped', 'ship quantity', 'qty', 'Ship Qty', 'Purchase History Ordered Quantity'])
        price_col = _pick(lower_to_orig, [
                          'invoice $', 'invoice$', 'invoice amount', 'invoice', 'price', 'Invoice $'])

        if not ndc_col or not ship_col or not price_col:
            raise ValueError(
                f"Vendor file '{os.path.basename(vp)}' must contain NDC/UPC, Ship Qty, and Invoice $ (or equivalents). "
                f"Found: {list(raw.columns)}"
            )

        # Optional date column (for Kinray "latest" logic)
        date_col = _pick(lower_to_orig, [
                         'invoice date', 'ship date', 'shipping date', 'order date', 'date'])
        ourcase_col = _pick(lower_to_orig, ['ourcase', 'our case', 'case',
                                            'invoice #', 'invoice number', 'inv #',
                                            'document number', 'doc #', 'doc no'])
        # # 🔹 NEW: optional drug-name column
        # drug_col = _pick(lower_to_orig, [
        #     'drug name', 'item description', 'description',
        #     'product name', 'item name','Description'
        # ])


        keep_cols = [ndc_col, ship_col, price_col] \
            + ([date_col] if date_col else []) \
            + ([ourcase_col] if ourcase_col else [])

        v = raw[keep_cols].copy()

        rename_map = {ndc_col: 'NDC #',
                      ship_col: 'Shipped', price_col: 'PRICE'}
        if date_col:
            rename_map[date_col] = 'DATE'
        if ourcase_col:
            rename_map[ourcase_col] = 'OURCASE'
        v.rename(columns=rename_map, inplace=True)
        # if drug_col:
        #     rename_map[drug_col] = 'Drug Name'   # 🔹 NEW
        # DATE
        if 'DATE' in v.columns:
            v['DATE'] = pd.to_datetime(v['DATE'], errors='coerce')
        else:
            v['DATE'] = pd.NaT

        # OURCASE: keep as string for stable sorting
        if 'OURCASE' in v.columns:
            v['OURCASE'] = v['OURCASE'].astype(str).str.strip()
        else:
            v['OURCASE'] = ''

        # Normalize NDC -> 11 digits
        v['NDC #'] = (v['NDC #'].astype(str)
                                .str.replace(r'\D', '', regex=True)
                                .str.strip()
                                .str.zfill(11))

        # Normalize Shipped -> numeric
        v['Shipped'] = (v['Shipped'].astype(str)
                        .str.replace(',', '', regex=False)
                        .str.replace('(', '-', regex=False)
                        .str.replace(')', '', regex=False)
                        .str.replace(r'[^0-9.\-]', '', regex=True)
                        .str.strip())
        v['Shipped'] = pd.to_numeric(v['Shipped'], errors='coerce').fillna(0)

        # Normalize PRICE -> numeric
        v['PRICE'] = (v['PRICE'].astype(str)
                                .str.replace(',', '', regex=False)
                                .str.replace('$', '', regex=False)
                                .str.replace('(', '-', regex=False)
                                .str.replace(')', '', regex=False)
                                .str.replace(r'[^0-9.\-]', '', regex=True)
                                .str.strip())
        v['PRICE'] = pd.to_numeric(v['PRICE'], errors='coerce').fillna(0)

        # Parse date (if present)
        if 'DATE' in v.columns:
            v['DATE'] = pd.to_datetime(v['DATE'], errors='coerce')
        else:
            v['DATE'] = pd.NaT

        # Vendor label (keep your convention)
        # ---------- ADD THIS BLOCK ----------
        # Derive a nice label from the file name (e.g. 'Kinray.xlsx' -> 'Kinray')
        vendor_label = os.path.splitext(os.path.basename(vp))[0]
        # If you want to force uppercase like 'MCK', you can do:
        # vendor_label = vendor_label.upper()

        v['Vendor'] = vendor_label
        # name = f'Vendor{i}'
        # v['Vendor'] = name
        # vendor_names.append(name)
        # Row-level unit price (only where Shipped > 0)
        v['__UnitPrice__'] = np.where(
            v['Shipped'] > 0,
            np.round(v['PRICE'] / v['Shipped'], 2),  # keep as float first
            np.nan
        )

        # remove infinities (e.g., division by 0)
        v['__UnitPrice__'] = np.where(np.isfinite(
            v['__UnitPrice__']), v['__UnitPrice__'], np.nan)

        # if you want whole numbers only, round and cast safely
        v['__UnitPrice__'] = v['__UnitPrice__'].round(0)  # no decimals
        v['__UnitPrice__'] = v['__UnitPrice__'].astype(
            'Int64')  # keeps NaN as <NA>

        is_kinray = 'kinray' in os.path.basename(vp).lower()
        if is_kinray:
            # keep all selectors we need to decide "latest"
            kinray_rows.append(
                v[['NDC #', 'DATE', 'OURCASE', 'Shipped',
                    'PRICE', '__UnitPrice__']].copy()
            )

        # Collect for aggregation/pivots
        vendor_frames_qty.append(v[['NDC #', 'Vendor', 'Shipped']])
        vendor_frames_price.append(v[['NDC #', 'Vendor', 'PRICE']])
        # --------- add this to feed Alternate NDC logic ----------
        cols_for_alt = ['NDC #', 'Shipped', 'PRICE', 'DATE', 'Vendor']
        if 'Drug Name' in v.columns:
            cols_for_alt.append('Drug Name')
        else:
            v['Drug Name'] = ""   # ensure column exists
            cols_for_alt.append('Drug Name')

        all_vendor_rows.append(v[cols_for_alt].copy())

    if all_vendor_rows:
        all_vendor_df = pd.concat(all_vendor_rows, ignore_index=True)
    else:
        all_vendor_df = pd.DataFrame(
            columns=['NDC #', 'Shipped', 'PRICE', 'DATE', 'Vendor'])

    #print("\n[DEBUG Vendor Preview]")
    #print(all_vendor_df[['NDC #', 'Drug Name']].head(), "\n")

    # ===== Qty pivot (sum of shipped by NDC×Vendor)
    vendor_combined = (pd.concat(vendor_frames_qty, ignore_index=True)
                       if vendor_frames_qty else pd.DataFrame(columns=['NDC #', 'Vendor', 'Shipped']))
    vendor_agg = vendor_combined.groupby(
        ['NDC #', 'Vendor'], as_index=False)['Shipped'].sum()
    vendor_pivot = (vendor_agg.pivot(index='NDC #', columns='Vendor', values='Shipped')
                    .fillna(0)
                    .reset_index())
    vendor_names = [c for c in vendor_pivot.columns if c != 'NDC #']
    # Capitalize vendor names
    vendor_rename_map = {vn: vn.upper() for vn in vendor_names}
    # Apply rename to pivots
    vendor_pivot = vendor_pivot.rename(columns=vendor_rename_map)
    # Update vendor_names list
    vendor_names = [vn.upper() for vn in vendor_names]

    # ===== Price pivot (last seen price by NDC×Vendor)
    vendor_price_combined = (pd.concat(vendor_frames_price, ignore_index=True)
                             if vendor_frames_price else pd.DataFrame(columns=['NDC #', 'Vendor', 'PRICE']))
    vendor_price_agg = (vendor_price_combined
                        .groupby(['NDC #', 'Vendor'], as_index=False)['PRICE']
                        .last())
    vendor_price_pivot = (vendor_price_agg.pivot(index='NDC #', columns='Vendor', values='PRICE')
                          .fillna(0)
                          .reset_index())
    vendor_price_pivot = vendor_price_pivot.rename(columns={c: (c.upper() + "_PRICE")
                                                            if c not in ("NDC #") else c for c in vendor_price_pivot.columns})
    vendor_price_pivot.columns = [
        'NDC #'] + [f'{c}_PRICE' for c in vendor_price_pivot.columns if c != 'NDC #']

    if kinray_rows:
        kinray_all = pd.concat(kinray_rows, ignore_index=True)

        # Treat missing DATE as very old, so real dates win
        min_ts = pd.Timestamp(1970, 1, 1)
        kinray_all['__DATE__'] = kinray_all['DATE'].fillna(min_ts)

        # If OURCASE is purely numeric in many files, try to rank; otherwise string compare works
        # Sort by NDC, then DATE asc, then OURCASE asc; keep last = latest
        kinray_latest = (
            kinray_all
            # must have calculable unit price
            .dropna(subset=['__UnitPrice__'])
            .sort_values(['NDC #', '__DATE__', 'OURCASE'])
            .drop_duplicates(subset=['NDC #'], keep='last')  # latest per NDC
            .loc[:, ['NDC #', '__UnitPrice__']]
            .rename(columns={'__UnitPrice__': 'Kinray_UPrice'})
        )
    else:
        kinray_all = pd.DataFrame(columns=['NDC #', 'DATE', 'OURCASE', 'Shipped', 'PRICE', '__UnitPrice__'])
        kinray_latest = pd.DataFrame(columns=['NDC #', 'Kinray_UPrice'])

    kinray_price_map = dict(
        zip(kinray_latest['NDC #'], kinray_latest['Kinray_UPrice']))

    # ===== END VENDOR AGGREGATION =====

    return vendor_pivot, vendor_price_pivot, vendor_names, kinray_all, kinray_latest, all_vendor_df
