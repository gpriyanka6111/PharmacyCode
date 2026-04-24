# Custom Log CSV parsing: status filtering, BIN normalization, insurance summary, and column renaming helpers.

import re

import numpy as np
import pandas as pd


def _normalize_status_value(v):
    s = '' if pd.isna(v) else str(v)
    return re.sub(r'[^a-z0-9]+', '', s.strip().lower())


def _filter_custom_log_transmitted_paid_ins(df):
    """
    Keep only rows whose status is Transmitted or Paid-Ins.
    Returns (filtered_df, status_col_name, kept_rows, dropped_rows).
    Raises ValueError if a usable status column cannot be identified.
    """
    if df is None or df.empty:
        return df.copy(), None, 0, 0

    col_lookup = {str(c).strip().lower(): c for c in df.columns}
    preferred = [
        'status', 'rx status', 'claim status', 'transaction status',
        'transmission status', 'payment status', 'rx state', 'state'
    ]

    status_col = None
    for key in preferred:
        if key in col_lookup:
            status_col = col_lookup[key]
            break

    allowed = {'transmitted', 'paidins', 'paidcash'}

    # If no direct status column match, infer by values.
    if status_col is None:
        for c in df.columns:
            vals = pd.Series(df[c]).dropna().astype(str)
            if vals.empty:
                continue
            norm_vals = set(vals.map(_normalize_status_value).unique().tolist())
            if ('transmitted' in norm_vals) and ('paidins' in norm_vals):
                status_col = c
                break

    if status_col is None:
        raise ValueError(
            "Custom Log must include a status column containing 'Transmitted' and 'Paid-Ins' values."
        )

    status_norm = df[status_col].map(_normalize_status_value)
    mask = status_norm.isin(allowed)
    kept = int(mask.sum())
    dropped = int((~mask).sum())
    return df.loc[mask].copy(), status_col, kept, dropped


def _build_insurance_summary(log_df, bin_df):
    """
    Returns:
      {
        "total_rx": <int>,  # unique Rx # (fallback to row count if missing)
        "by_processor": [
            {"processor": "CAREMARK", "rx_count": 200, "total_paid": 200000.00},
            ...
        ],
        "processors": ["CAREMARK","OPTUMRX",...]
      }
    """
    df = log_df.copy()

    # normalize numerics
    for c in ['Ins Paid Plan 1', 'Ins Paid Plan 2']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # pick winner per row
    df['Winning_BIN'] = np.where(df.get('Ins Paid Plan 1', 0) >= df.get('Ins Paid Plan 2', 0),
                                 df.get('Plan 1 BIN', ''), df.get('Plan 2 BIN', ''))
    df['Winning_BIN'] = df['Winning_BIN'].astype(
        str).str.replace(r'\D', '', regex=True).str.zfill(6)

    df['Winning_Paid'] = np.where(
        df['Winning_BIN'] == df.get('Plan 1 BIN', ''),
        df.get('Ins Paid Plan 1', 0),
        df.get('Ins Paid Plan 2', 0)
    )

    # BIN -> Processor
    bin_df = bin_df.copy()
    bin_df['BIN'] = bin_df['BIN'].astype(str).str.replace(
        r'\D', '', regex=True).str.zfill(6)
    bin_df['Processor'] = bin_df['Processor'].astype(str).str.strip()
    bin_to_proc = dict(zip(bin_df['BIN'], bin_df['Processor']))

    df['Processor'] = df['Winning_BIN'].map(bin_to_proc)

    # RX count
    if 'Rx #' in df.columns:
        total_rx = df['Rx #'].astype(str).str.strip().replace(
            '', np.nan).dropna().nunique()
    else:
        total_rx = len(df)

    # group by processor
    grp = (df.dropna(subset=['Processor'])
             .groupby('Processor', as_index=False)
             .agg(rx_count=('Winning_BIN', 'count'),
                  total_paid=('Winning_Paid', 'sum')))

    # nice ordering by total $
    grp = grp.sort_values('total_paid', ascending=False)
    processors = grp['Processor'].tolist()

    by_processor = [
        {
            "processor": r['Processor'],
            "rx_count": int(r['rx_count']),
            "total_paid": float(r['total_paid'])
        }
        for _, r in grp.iterrows()
    ]

    return {
        "total_rx": int(total_rx),
        "by_processor": by_processor,
        "processors": processors
    }
