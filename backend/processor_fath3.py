"""
Fath3 Processor - adapted from new/test_fath3.py

Notes:
- The Electron UI shows a fixed 4-column results table (Ref, Stock_Qty, Calc_Mov_Qty, Difference).
- The original script compares per-localisation; here we aggregate across the relevant localisations
  for each movement file key to fit the UI output.
"""

import os
import re
import pandas as pd


# Sheet arguments configuration
sheet_args = {
    'pet': {
        'file_keyword': 'pet',
        'sheet_name': ['MOUV PARC03-2023'],
        'header_must_contain': ['Date', 'LOCALITATION', 'PRODOUITE', 'Q ST PV'],
        'mov cols': ['PRODOUITE', 'Q ST PV'],
        'mov localisation cols': ['LOCALITATION', 'LOCALISATION'],
        'exclude_localisations': ['ATT TRIAGE'],
        'stock ref cols': ['RF'],
        'stock qty cols': ['S REEL', 'S REEL ', 'S RÉEL'],
        'stock localisation cols': ['LOCALISATION', 'LOCALISATION '],
    },
    'triage': {
        'file_keyword': 'triage',
        'sheet_name': ['MOUVM'],
        'header_must_contain': ['DATE', 'PRODUIT', 'STOCK'],
        'mov cols': ['PRODUIT', 'STOCK'],
        'localisation': ['ATT TRIAGE'],
        'no_localisation_column': True,
    },
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'localisation': ['LOCALITATION', 'LOCALISATION', 'LOCALISATION '],
    'ref': ['PRODOUITE', 'PRODUIT', 'REF', 'REFERENCE', 'REFERENCE ', 'Référence', 'REF PRODUIT'],
    'quantity': ['Q ST PV', 'Q ST PV ', 'Quantité', 'STOCK PV', 'STOCK'],
}

stock_possible_col_names = {
    'ref': ['RF', 'RF ', 'REF', 'REFERENCE', 'REFERENCE '],
    'quantity': ['S REEL', 'S REEL ', 'S RÉEL', 'QUANTITE', 'QUANTITÉ'],
    'localisation': ['LOCALISATION', 'LOCALISATION '],
}


def get_ateliers():
    return list(sheet_args.keys())


def _norm_text(val: object) -> str:
    if val is None:
        return ''
    return str(val).strip()


def _find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    norm_map = {_norm_text(col).casefold(): col for col in cols}
    for c in candidates:
        key = _norm_text(c).casefold()
        if key in norm_map:
            return norm_map[key]
    return None


def _detect_header_row(excel_path: str, sheet_name, must_contain: list[str], max_rows: int = 80) -> int:
    tmp = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    scan_rows = min(max_rows, len(tmp))
    for idx in range(scan_rows):
        row = tmp.iloc[idx]
        row_values = [_norm_text(v) for v in row.values if pd.notna(v)]
        if not row_values:
            continue
        row_set = set(row_values)
        hits = sum(1 for k in must_contain if k in row_set)
        if hits >= 2:
            return int(idx)
        if any('date' in v.casefold() for v in row_values):
            return int(idx)
    return 0


def _coerce_date(series: pd.Series) -> pd.Series:
    if series is None:
        return series

    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    if pd.api.types.is_numeric_dtype(series):
        s_num = pd.to_numeric(series, errors='coerce')
        median = s_num.dropna().median() if s_num.notna().any() else None
        if median is not None and 20000 <= median <= 60000:
            return pd.to_datetime(s_num, origin='1899-12-30', unit='D', errors='coerce')

    s_str = series.astype('string').str.strip()
    s_str = s_str.replace({'': pd.NA})
    return pd.to_datetime(s_str, errors='coerce', dayfirst=True)


def _normalize_ref(series: pd.Series) -> pd.Series:
    s = series.astype('string').str.strip()
    return s.str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)


def _coerce_qty(series: pd.Series) -> pd.Series:
    s_str = series.astype('string')
    s_str = s_str.str.replace('\u00A0', ' ', regex=False)
    s_str = s_str.str.replace(' ', '', regex=False)
    s_str = s_str.str.replace(',', '.', regex=False)
    return pd.to_numeric(s_str, errors='coerce').fillna(0.0)


def load_stock(stock_file_path: str, stock_sheet_name: str = 'STOCKS') -> tuple[pd.DataFrame, str, str, str]:
    header_idx = _detect_header_row(
        stock_file_path,
        stock_sheet_name,
        must_contain=['RF', 'S REEL', 'LOCALISATION'],
    )
    stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=header_idx)
    stock = stock.dropna(how='all')

    stock_ref_col = _find_column(stock, stock_possible_col_names['ref'])
    stock_qty_col = _find_column(stock, stock_possible_col_names['quantity'])
    stock_loc_col = _find_column(stock, stock_possible_col_names['localisation'])

    if not stock_ref_col or not stock_qty_col or not stock_loc_col:
        raise KeyError(
            f"Stock file is missing required columns. Found ref={stock_ref_col}, qty={stock_qty_col}, loc={stock_loc_col}. "
            f"Available: {stock.columns.tolist()}"
        )

    for col in stock.select_dtypes(include=['object']).columns:
        stock[col] = stock[col].astype('string').str.strip()

    stock[stock_ref_col] = _normalize_ref(stock[stock_ref_col])
    stock[stock_qty_col] = _coerce_qty(stock[stock_qty_col]).round(4)

    # Drop rows missing key fields
    stock = stock.dropna(subset=[stock_ref_col, stock_loc_col])
    stock = stock[stock[stock_ref_col].astype('string').str.strip() != '']
    stock = stock[stock[stock_loc_col].astype('string').str.strip() != '']

    return stock, stock_ref_col, stock_qty_col, stock_loc_col


def _read_mov(mov_file_path: str, possible_sheets: list, header_must_contain: list[str]) -> tuple[pd.DataFrame, object]:
    mov = None
    used_sheet = None

    for sheet in possible_sheets:
        try:
            header_idx = _detect_header_row(mov_file_path, sheet, must_contain=header_must_contain)
            mov = pd.read_excel(mov_file_path, sheet_name=sheet, header=header_idx)
            used_sheet = sheet
            break
        except ValueError:
            continue
        except Exception:
            continue

    if mov is None:
        raise ValueError(f"Could not read any of sheets {possible_sheets} from {mov_file_path}")

    mov = mov.dropna(how='all')
    return mov, used_sheet


def process_atelier(atelier_key: str, stock_df: pd.DataFrame, stock_ref_col: str, stock_qty_col: str, stock_loc_col: str,
                   mov_file_path: str, month: str, overrides: dict | None = None) -> dict:
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}

    args = sheet_args[atelier_key]
    overrides = overrides or {}

    sheet_override = overrides.get('sheetName')
    ref_override = overrides.get('refCol')
    qty_override = overrides.get('qtyCol')

    possible_sheets = args.get('sheet_name', [])
    if not isinstance(possible_sheets, list):
        possible_sheets = [possible_sheets]
    if sheet_override:
        possible_sheets = [sheet_override]

    header_must_contain = args.get('header_must_contain') or ['Date', 'LOCALITATION', 'PRODOUITE', 'Q ST PV']

    try:
        mov, _ = _read_mov(mov_file_path, possible_sheets, header_must_contain)

        # Resolve columns
        mov_cols_override = args.get('mov cols')
        if isinstance(mov_cols_override, (list, tuple)) and len(mov_cols_override) == 2:
            default_ref_col, default_qty_col = mov_cols_override
        else:
            default_ref_col = _find_column(mov, mov_possible_col_names['ref'])
            default_qty_col = _find_column(mov, mov_possible_col_names['quantity'])

        ref_col = ref_override or default_ref_col
        qty_col = qty_override or default_qty_col

        no_loc_col = bool(args.get('no_localisation_column', False))
        mov_loc_col = None if no_loc_col else _find_column(mov, args.get('mov localisation cols', mov_possible_col_names['localisation']))
        mov_date_col = _find_column(mov, mov_possible_col_names['date'])

        if not ref_col or not qty_col or not mov_date_col or (not no_loc_col and not mov_loc_col):
            return {
                'error': f"Missing required movement columns. Found ref={ref_col}, qty={qty_col}, loc={mov_loc_col}, date={mov_date_col}. Available: {mov.columns.tolist()}",
                'matches': [],
                'discrepancies': []
            }

        # Clean & normalize
        for col in mov.select_dtypes(include=['object']).columns:
            mov[col] = mov[col].astype('string').str.strip()

        mov[mov_date_col] = _coerce_date(mov[mov_date_col])
        if no_loc_col:
            mov = mov.dropna(subset=[mov_date_col])
        else:
            mov = mov.dropna(subset=[mov_date_col, mov_loc_col])
            mov = mov[mov[mov_loc_col].astype('string').str.strip() != '']

        # Filter by Date (keep months <= target month for 2025; keep previous years)
        target_month = int(month)
        target_year = 2025
        mask = (mov[mov_date_col].dt.year < target_year) | ((mov[mov_date_col].dt.year == target_year) & (mov[mov_date_col].dt.month <= target_month))
        mov = mov[mask]

        mov[ref_col] = _normalize_ref(mov[ref_col])
        mov[qty_col] = _coerce_qty(mov[qty_col])

        # Localisation selection
        exclude_locs = [str(x).strip() for x in (args.get('exclude_localisations') or []) if str(x).strip()]
        include_locs = args.get('localisation')

        if include_locs:
            include_locs = [str(x).strip() for x in include_locs if str(x).strip()]
        else:
            include_locs = sorted(stock_df[stock_loc_col].astype('string').str.strip().unique().tolist())

        if exclude_locs:
            exclude_set = set(exclude_locs)
            include_locs = [l for l in include_locs if l not in exclude_set]

        # Filter stock
        stock_filtered = stock_df[stock_df[stock_loc_col].astype('string').str.strip().isin(include_locs)].copy()

        # Filter movement (if it has a localisation column)
        if not no_loc_col:
            if exclude_locs:
                mov = mov[~mov[mov_loc_col].astype('string').str.strip().isin(exclude_locs)]
            mov_filtered = mov[mov[mov_loc_col].astype('string').str.strip().isin(include_locs)].copy()
        else:
            mov_filtered = mov.copy()

        # Aggregate
        stock_agg = stock_filtered.groupby(stock_ref_col)[stock_qty_col].sum().reset_index()
        stock_agg.rename(columns={stock_ref_col: 'Ref', stock_qty_col: 'Stock_Qty'}, inplace=True)
        stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)

        mov_agg = mov_filtered.groupby(ref_col)[qty_col].sum().reset_index()
        mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
        mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)

        comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
        comparison_df['Difference'] = (comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']).round(2)

        discrepancies = comparison_df[comparison_df['Difference'].abs() > 0.02].sort_values(by='Difference', ascending=False)
        matches = comparison_df[comparison_df['Difference'].abs() <= 0.02]

        return {
            'matches': matches.to_dict('records'),
            'discrepancies': discrepancies.to_dict('records')
        }

    except Exception as e:
        return {'error': str(e), 'matches': [], 'discrepancies': []}


def process_all(stock_file, matched_files, month):
    results = {}
    try:
        stock_df, stock_ref_col, stock_qty_col, stock_loc_col = load_stock(stock_file['path'], stock_sheet_name='STOCKS')
    except Exception as e:
        return {'_error': f'Failed to load stock file: {str(e)}'}

    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(
            atelier_key,
            stock_df,
            stock_ref_col,
            stock_qty_col,
            stock_loc_col,
            mov_file['path'],
            month,
            overrides=None,
        )

    return results


def process_all_with_overrides(stock_file, matched_files, month, overrides):
    results = {}
    try:
        stock_df, stock_ref_col, stock_qty_col, stock_loc_col = load_stock(stock_file['path'], stock_sheet_name='STOCKS')
    except Exception as e:
        return {'_error': f'Failed to load stock file: {str(e)}'}

    for atelier_key, mov_file in matched_files.items():
        atelier_overrides = overrides.get(atelier_key, {}) if overrides else {}
        results[atelier_key] = process_atelier(
            atelier_key,
            stock_df,
            stock_ref_col,
            stock_qty_col,
            stock_loc_col,
            mov_file['path'],
            month,
            overrides=atelier_overrides,
        )

    return results
