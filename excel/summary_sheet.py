# Builds the "Summary" sheet: processor-level Insurance $, Purchased $, Net, and Order Estimate rows.

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def add_summary_sheet(
    wb,
    processed_source="Processed Data",           # str sheet name OR a worksheet
    needs_title="Needs to be ordered - All",
    header_row=3,
    data_start_row=4,
    pharmacy_name=None,
    date_range=None,
):
    """
    Builds a 'Summary' sheet with columns = processors (+ ALL_PBM if present) and rows:
      - Insurance $ Paid
      - $$ Purchased (Kinray)
      - Net (Paid − Purchased)
      - Insurance-wise Order Estimate ($)  <-- pulled from Needs sheet

    Notes:
    - processed_source can be a sheet name (str) or an openpyxl Worksheet.
    - Formats money as currency.
    """

    # --- Resolve processed worksheet ---
    if isinstance(processed_source, str):
        if processed_source not in wb.sheetnames:
            return
        ws_pd = wb[processed_source]
        processed_title = processed_source
    else:
        # assume it's a worksheet
        ws_pd = processed_source
        processed_title = ws_pd.title

    # --- read headers from the header_row ---
    headers = [
        ws_pd.cell(row=header_row, column=c).value
        for c in range(1, ws_pd.max_column + 1)
    ]
    headers = [h for h in headers if h]

    def _procs_by_suffix(suffix: str):
        return sorted({
            h[:-len(suffix)]
            for h in headers
            if isinstance(h, str) and h.endswith(suffix)
        })

    procs_T = _procs_by_suffix("_T")
    procs_Pur = _procs_by_suffix("_Pur")
    procs_Net = _procs_by_suffix("_Net")

    processors = sorted(set(procs_T) | set(procs_Pur) | set(procs_Net))

    # include ALL_PBM if any ALL_PBM_* exists
    if "ALL_PBM" not in processors and any(
        isinstance(h, str) and h.startswith("ALL_PBM_") for h in headers
    ):
        processors.append("ALL_PBM")

    if not processors:
        # create a small Summary sheet indicating no processor columns
        if "Summary" in wb.sheetnames:
            del wb["Summary"]
        ws = wb.create_sheet("Summary")
        ws["A1"] = "No processor metric columns (_T, _Pur, _Net) found in processed sheet."
        return

    # helper: find exact header col index
    def col_idx_for(header_text):
        for c in range(1, ws_pd.max_column + 1):
            if ws_pd.cell(row=header_row, column=c).value == header_text:
                return c
        return None

    # build processor -> column letter maps for each band
    def band_cols(suffix: str):
        out = {}
        for p in processors:
            hdr = f"{p}{suffix}"
            idx = col_idx_for(hdr)
            if idx:
                out[p] = get_column_letter(idx)
        return out

    cols_T = band_cols("_T")
    cols_Pur = band_cols("_Pur")
    cols_Net = band_cols("_Net")

    # find last data row (use max_row as a safe upper bound)
    last_data_row = ws_pd.max_row

    # --- Create/replace Summary at last index ---
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws = wb.create_sheet("Summary")

    # Title (row 1)
    end_col = len(processors) + 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    title_txt = "Summary"
    ws.cell(row=1, column=1, value=title_txt).font = Font(bold=True, size=16)
    ws.cell(row=1, column=1).alignment = Alignment(
        horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Optional subtitle with pharmacy/date (row 2) if provided
    if pharmacy_name or date_range:
        ws.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=end_col)
        if pharmacy_name and date_range:
            sub = f"Summary of {pharmacy_name} for the date range {date_range}"
        else:
            sub = " · ".join([t for t in [pharmacy_name, date_range] if t])
        ws.cell(row=2, column=1, value=sub).alignment = Alignment(
            horizontal="center")
        ws.row_dimensions[2].height = 22
        header_base_row = 3
    else:
        header_base_row = 2

    # Header row (Metric + processors)
    ws.cell(row=header_base_row, column=1,
            value="Metric").font = Font(bold=True)
    for j, p in enumerate(processors, start=2):
        cell = ws.cell(row=header_base_row, column=j, value=p)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Row labels start below header
    start_data_row = header_base_row + 1
    metrics = [
        "Insurance $(BestRX)",
        "100% $$ Purchased (Kinray)",
        "Net (Paid − Purchased)",
        "Needs to Ordered Sheet, Insurance-wise Order Estimate ($)",
    ]

    for i, m in enumerate(metrics):
        cell = ws.cell(row=start_data_row + i, column=1, value=m)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Increase row height for better visibility
        ws.row_dimensions[start_data_row + i].height = 28

    # Currency format
    money_fmt = '"$"#,##0.00'

    # Fill formulas for the first 3 metrics from Processed Data
    row_paid = start_data_row
    row_pur = start_data_row + 1
    row_net = start_data_row + 2
    row_est = start_data_row + 3

    for j, p in enumerate(processors, start=2):
        # Paid (T)
        if p in cols_T:
            col = cols_T[p]
            ws.cell(row=row_paid, column=j,
                    value=f"=SUM('{processed_title}'!{col}{data_start_row}:{col}{last_data_row})"
                    ).number_format = money_fmt
        else:
            ws.cell(row=row_paid, column=j, value=0).number_format = money_fmt

        # Purchased (Pur)
        if p in cols_Pur:
            col = cols_Pur[p]
            ws.cell(row=row_pur, column=j,
                    value=f"=SUM('{processed_title}'!{col}{data_start_row}:{col}{last_data_row})"
                    ).number_format = money_fmt
        else:
            ws.cell(row=row_pur, column=j, value=0).number_format = money_fmt

        # Net
        if p in cols_Net:
            col = cols_Net[p]
            ws.cell(row=row_net, column=j,
                    value=f"=SUM('{processed_title}'!{col}{data_start_row}:{col}{last_data_row})"
                    ).number_format = money_fmt
        else:
            ws.cell(row=row_net, column=j, value=0).number_format = money_fmt

    # --- Insurance-wise Order Estimate ($) from Needs sheet
    if needs_title in wb.sheetnames:
        ws_need = wb[needs_title]

        def find_header_col(ws0, text, hdr_row=2):
            for c in range(1, ws0.max_column + 1):
                if (ws0.cell(row=hdr_row, column=c).value or "") == text:
                    return c
            return None

        label_col_idx = find_header_col(
            ws_need, "Insurance-wise Order Estimate ($)")
        value_col_idx = find_header_col(ws_need, "Amount")

        # Fallback if headers differ: default to first two columns
        if not label_col_idx:
            label_col_idx = 1
        if not value_col_idx:
            value_col_idx = 2

        label_col_letter = get_column_letter(label_col_idx)
        value_col_letter = get_column_letter(value_col_idx)
        last_need_row = ws_need.max_row

        for j, p in enumerate(processors, start=2):
            # Looks up "<processor>_D" label in Needs sheet
            formula = (
                f"=IFERROR(INDEX('{needs_title}'!${value_col_letter}$3:${value_col_letter}${last_need_row},"
                f"MATCH(\"{p}_D\", '{needs_title}'!${label_col_letter}$3:${label_col_letter}${last_need_row}, 0)), 0)"
            )
            c = ws.cell(row=row_est, column=j, value=formula)
            c.number_format = money_fmt
    else:
        # No Needs sheet -> zeroes
        for j, _ in enumerate(processors, start=2):
            ws.cell(row=row_est, column=j, value=0).number_format = money_fmt

    # Freeze & filter
    freeze_row = header_base_row + 1
    ws.freeze_panes = f"B{freeze_row}"
    ws.auto_filter.ref = f"A{header_base_row}:{get_column_letter(ws.max_column)}{ws.max_row}"

    # Column widths
    ws.column_dimensions['A'].width = 32
    for j in range(2, len(processors) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 16

    # Borders & header fill
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    for r in range(header_base_row, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).border = thin

    header_fill = PatternFill(start_color="D0CECE",
                              end_color="D0CECE", fill_type="solid")
    for c in range(1, ws.max_column + 1):
        ws.cell(row=header_base_row, column=c).fill = header_fill
