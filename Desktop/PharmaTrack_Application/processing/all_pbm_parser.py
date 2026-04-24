# Drug Usage Summary (ALL PBM) CSV parser: auto-delimiter detection, column mapping, NDC normalization.

import re

import pandas as pd


def _load_all_pbm_csv(path):
    """
    Read ALL PBM file (CSV/TSV). Auto-detect delimiter.
    Returns: ['NDC #','ALL_PBM_Q','ALL_PBM_T','ALL_PBM_DrugName']
    (No aggregation: if an NDC appears multiple times, we keep the LAST row.)
    """
    # Try auto delimiter; handle BOM
    df = pd.read_csv(path, dtype=str, encoding='utf-8-sig',
                     sep=None, engine='python')
    if df.shape[1] == 1:
        # Often actually TAB-delimited saved as .csv
        try:
            df = pd.read_csv(path, dtype=str, encoding='utf-8-sig', sep='\t')
        except Exception:
            pass

    # Normalize headers -> lowercase, collapse spaces, remove nbsp
    def norm(s):
        return re.sub(r'\s+', ' ', str(s).replace('\u00A0', ' ')).strip().lower()
    df.columns = [norm(c) for c in df.columns]

    # Column pickers (exact first, then contains)
    def pick(*cands):
        for k in cands:
            if k in df.columns:
                return k
        for k in cands:
            for col in df.columns:
                if k in col:
                    return col
        return None

    ndc_k = pick('ndc #', 'ndc#', 'ndc', 'ndc upc', 'ndc/upc')
    qty_k = pick('quantity', 'qty', 'total quantity', 'total qty', 'Quantity')
    total_k = pick('all_pbm_t', 'total $', 'total$',
                   'total amount', 'amount', 'total', 'Total')
    name_k = pick('drug name', 'drug', 'name')

    if ndc_k is None:
        # Return empty frame with expected columns so merge won't break
        return pd.DataFrame(columns=['NDC #', 'ALL_PBM_Q', 'ALL_PBM_T', 'ALL_PBM_DrugName'])

    out = pd.DataFrame()
    out['NDC #'] = (df[ndc_k].astype(str)
                    .str.replace(r'\D', '', regex=True)
                    .str.zfill(11))

    # ALL_PBM_Q (optional; default 0)
    if qty_k:
        q = (df[qty_k].astype(str)
             .str.replace(',', '', regex=False)
             .str.replace(r'[^0-9.\-]', '', regex=True))
        out['ALL_PBM_Q'] = pd.to_numeric(q, errors='coerce').fillna(0)
    else:
        out['ALL_PBM_Q'] = 0

    # ALL_PBM_T (Total $) — verbatim per row, no aggregation
    if total_k:
        t = (df[total_k].astype(str)
             .str.replace(',', '', regex=False)
             .str.replace('$', '', regex=False)
             .str.replace('(', '-', regex=False)
             .str.replace(')', '', regex=False)
             .str.replace(r'[^0-9.\-]', '', regex=True))
        out['ALL_PBM_T'] = pd.to_numeric(t, errors='coerce').fillna(0)
    else:
        out['ALL_PBM_T'] = 0

    out['ALL_PBM_DrugName'] = df[name_k].astype(
        str).str.strip() if name_k else pd.NA

    # No aggregation: if duplicates exist, keep the LAST one from the file
    out = out.drop_duplicates(subset=['NDC #'], keep='last')

    return out[['NDC #', 'ALL_PBM_Q', 'ALL_PBM_T', 'ALL_PBM_DrugName']]
