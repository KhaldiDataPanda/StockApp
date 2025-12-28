"""Mdoukal Processor - adapted from new/test_mdoukal.py"""

import pandas as pd


sheet_args = {
    "couture femmes": {
        "sheet_name": ["STC"],
        "localisation": ["ATELIER COUTURE FEMME"],
        "mov cols": ["REFERENCE", "STOCK"],
    },
    "orielles": {
        "sheet_name": ["Sheet1"],
        "localisation": ["ATELIER PREPARATION ORIELLES FINI"],
        "mov cols": ["REFERENCE", "STOCK"],
    },
    "magasin": {
        "sheet_name": ["MOV"],
        "localisation": ["MAGASIN PRINCIPAL UNITE M'DOUKEL"],
        "mov cols": ["REFERENCE", "STOCK"],
    },
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'ref': ['Row Labels', 'Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'REF', 'REF PRODUIT', 'Référence', 'referance', 'البيان', 'المرجع', 'نوع', 'REFERNCE', 'التعيين'],
    'quantity': ['Quantité', 'STOCK PV', 'STOCK FIBRE ', 'STOCK', 'STOCK ', 'STOCKS', 'ST-P', 'STOKS', 'STOCK U', 'المخزون', 'الكمية', 'العدد', 'STOCK/M', 'STOCK POIDS', 'باقي', 'Q-STOCKS'],
}


def get_ateliers():
    return list(sheet_args.keys())


def _find_header_row_by_date(excel_path: str, sheet_name, date_candidates: list[str], max_rows: int = 80) -> int:
    tmp = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    scan_rows = min(max_rows, len(tmp))
    for idx in range(scan_rows):
        row = tmp.iloc[idx]
        row_values = [str(v).strip() for v in row.values if pd.notna(v)]
        if any(name in row_values for name in date_candidates):
            return int(idx)
    return 0


def _normalize_ref(series: pd.Series) -> pd.Series:
    series = series.astype('string').str.strip()
    return series.str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)


def load_stock(stock_file_path: str):
    # Try to read MOUV, fallback to first sheet
    xl = pd.ExcelFile(stock_file_path)
    sheet_to_use = 'MOUV' if 'MOUV' in xl.sheet_names else xl.sheet_names[0]

    header_idx = _find_header_row_by_date(stock_file_path, sheet_to_use, mov_possible_col_names['date'])
    stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=header_idx)
    stock = stock.dropna(how='all')

    # Expect REF/QUANTITE/LOCALISATION, but be a bit tolerant
    ref_col = 'REF' if 'REF' in stock.columns else next((c for c in stock.columns if 'ref' in str(c).lower()), None)
    qty_col = 'QUANTITE' if 'QUANTITE' in stock.columns else next((c for c in stock.columns if 'quant' in str(c).lower()), None)
    loc_col = 'LOCALISATION' if 'LOCALISATION' in stock.columns else next((c for c in stock.columns if 'local' in str(c).lower()), None)

    if not ref_col or not qty_col or not loc_col:
        raise KeyError(f"Stock file missing columns. Found ref={ref_col}, qty={qty_col}, loc={loc_col}. Available: {stock.columns.tolist()}")

    for col in stock.select_dtypes(include=['object']).columns:
        stock[col] = stock[col].astype('string').str.strip()

    stock[ref_col] = _normalize_ref(stock[ref_col])
    stock[qty_col] = pd.to_numeric(stock[qty_col], errors='coerce').fillna(0.0).round(2)

    return stock, ref_col, qty_col, loc_col


def _read_mov(mov_file_path: str, possible_sheets: list):
    last_error = None
    for sheet in possible_sheets:
        try:
            header_idx = _find_header_row_by_date(mov_file_path, sheet, mov_possible_col_names['date'])
            mov = pd.read_excel(mov_file_path, sheet_name=sheet, header=header_idx if header_idx else None)
            return mov, sheet
        except ValueError as e:
            last_error = e
            continue
        except Exception as e:
            last_error = e
            continue
    raise RuntimeError(f"Could not read any of sheets {possible_sheets} from {mov_file_path}. Last error: {last_error}")


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

    try:
        mov, _ = _read_mov(mov_file_path, possible_sheets)
        mov = mov.dropna(how='all')

        # Find columns
        found_cols = {}
        for col_type, possible_names in mov_possible_col_names.items():
            for name in possible_names:
                if name in mov.columns:
                    found_cols[col_type] = name
                    break

        mov_cols_override = args.get('mov cols')
        default_ref = mov_cols_override[0] if isinstance(mov_cols_override, (list, tuple)) and len(mov_cols_override) == 2 else found_cols.get('ref')
        default_qty = mov_cols_override[1] if isinstance(mov_cols_override, (list, tuple)) and len(mov_cols_override) == 2 else found_cols.get('quantity')

        ref_col = ref_override or default_ref
        qty_col = qty_override or default_qty

        if not ref_col or not qty_col:
            return {'error': f"Could not find ref/quantity columns. Available: {mov.columns.tolist()}", 'matches': [], 'discrepancies': []}

        # Date filter (months <= target for 2025; keep previous years)
        if 'date' in found_cols:
            date_col = found_cols['date']
            mov[date_col] = pd.to_datetime(mov[date_col], errors='coerce')
            target_month = int(month)
            target_year = 2025
            mask = (mov[date_col].dt.year < target_year) | ((mov[date_col].dt.year == target_year) & (mov[date_col].dt.month <= target_month))
            mov = mov[mask]

        for col in mov.select_dtypes(include=['object']).columns:
            mov[col] = mov[col].astype('string').str.strip()

        mov[ref_col] = _normalize_ref(mov[ref_col])
        mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce').fillna(0.0).round(2)

        mov_agg = mov.groupby(ref_col)[qty_col].sum().reset_index()
        mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
        mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)

        # Filter stock by localisation
        localisations = args.get('localisation') or []
        stock_filtered = stock_df[stock_df[stock_loc_col].isin(localisations)].copy()

        stock_agg = stock_filtered.groupby(stock_ref_col)[stock_qty_col].sum().reset_index()
        stock_agg.rename(columns={stock_ref_col: 'Ref', stock_qty_col: 'Stock_Qty'}, inplace=True)
        stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)

        comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
        comparison_df['Difference'] = (comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']).round(2)

        discrepancies = comparison_df[comparison_df['Difference'].abs() > 0.02].sort_values(by='Difference', ascending=False)
        matches = comparison_df[comparison_df['Difference'].abs() <= 0.02]

        return {'matches': matches.to_dict('records'), 'discrepancies': discrepancies.to_dict('records')}

    except Exception as e:
        return {'error': str(e), 'matches': [], 'discrepancies': []}


def process_all(stock_file, matched_files, month):
    results = {}
    try:
        stock_df, stock_ref_col, stock_qty_col, stock_loc_col = load_stock(stock_file['path'])
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
        stock_df, stock_ref_col, stock_qty_col, stock_loc_col = load_stock(stock_file['path'])
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
