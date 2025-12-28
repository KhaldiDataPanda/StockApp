"""Mags Processor - adapted from new/test_mags.py

This unit expects end-of-month stock to equal:
  previous_month_stock + movements_of_selected_month

To keep the existing UI (4-column results), we expose:
- Stock_Qty: current stock quantity
- Calc_Mov_Qty: expected end quantity (Prev_Stock_Qty + Mov_Qty)
- Difference: Stock_Qty - Calc_Mov_Qty
"""

import os
import re
import pandas as pd


sheet_args = {
    'magz': {
        'sheet_name': ['MOUV', 'MOV', 'MAG', 0],
    }
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'ref': ['REFERENCE', 'REFERANCE', 'REFERENCe', 'Référence', 'REF', 'REF PRODUIT', 'Row Labels', 'Référence\nFournisseur', 'reference', 'referance', 'REFERNCE', 'البيان', 'المرجع', 'نوع', 'التعيين'],
    'quantity': ['QUANTITE', 'Quantité', 'QTE', 'QTY', 'STOCK', 'STOCKS', 'STOCK ', 'ST-P', 'STOKS', 'STOCK U', 'STOCK/M', 'STOCK POIDS', 'STOCK PV', 'STOCK FIBRE ', 'باقي', 'Q-STOCKS', 'المخزون', 'الكمية', 'العدد'],
}


def get_ateliers():
    return list(sheet_args.keys())


def prev_month_year(month_str: str, year: int) -> tuple[str, int]:
    m = int(month_str)
    if m == 1:
        return '12', year - 1
    return f"{m - 1:02d}", year


def _normalize_ref(series: pd.Series) -> pd.Series:
    series = series.astype('string').str.strip()
    return series.str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)


def _read_excel_with_sheet_fallback(excel_path: str, sheet_candidates: list, header: int | None = None) -> tuple[pd.DataFrame, object]:
    last_error = None
    for sheet in sheet_candidates:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet, header=header)
            return df, sheet
        except ValueError as e:
            last_error = e
            continue
        except Exception as e:
            last_error = e
            continue
    raise RuntimeError(f"Could not read any of sheets {sheet_candidates} from {excel_path}. Last error: {last_error}")


def _find_header_row(excel_path: str, sheet_name, possible_names: list[str], max_scan_rows: int = 50) -> int:
    temp_df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    scan_rows = min(max_scan_rows, len(temp_df))
    for idx in range(scan_rows):
        row_values = [str(val).strip() for val in temp_df.iloc[idx].values if pd.notna(val)]
        if any(name in row_values for name in possible_names):
            return idx
    return 0


def _infer_year_from_stock_filename(path: str) -> int | None:
    name = os.path.basename(path)
    # Matches: STOCK 11-2025.xlsx, STOCK 11-2025 (1).xlsx, etc.
    m = re.search(r'(?:stock)\s*\d{1,2}\s*[-_](\d{4})', name, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def _find_neighbor_stock_file(current_stock_path: str, month: str, year: int) -> str | None:
    directory = os.path.dirname(current_stock_path)
    if not os.path.isdir(directory):
        return None

    expected1 = os.path.join(directory, f"STOCK {month}-{year}.xlsx")
    expected2 = os.path.join(directory, f"STOCK {month}-{year}.xls")

    if os.path.exists(expected1):
        return expected1
    if os.path.exists(expected2):
        return expected2

    # Fallback: scan directory for a file containing 'stock' and the month-year
    month_year_key = f"{int(month)}-{year}"
    for fname in os.listdir(directory):
        f = fname.lower()
        if not (f.endswith('.xlsx') or f.endswith('.xls')):
            continue
        if 'stock' in f and month_year_key in f.replace(' ', ''):
            return os.path.join(directory, fname)
    return None


def load_stock(stock_path: str) -> tuple[pd.DataFrame, str, str]:
    sheet_candidates = ['MOUV', 'MOV', 'MAG', 0]

    tmp, sheet_used = _read_excel_with_sheet_fallback(stock_path, sheet_candidates, header=None)
    header_candidates = mov_possible_col_names['ref'] + mov_possible_col_names['quantity']
    header_idx = 0
    try:
        header_idx = _find_header_row(stock_path, sheet_used, header_candidates)
    except Exception:
        header_idx = 0

    stock, _ = _read_excel_with_sheet_fallback(stock_path, [sheet_used], header=header_idx)
    stock = stock.dropna(how='all')

    found_cols = {}
    for col_type, possible_names in mov_possible_col_names.items():
        for name in possible_names:
            if name in stock.columns:
                found_cols[col_type] = name
                break

    if 'ref' not in found_cols or 'quantity' not in found_cols:
        raise KeyError(f"Could not find required columns in stock file. Columns: {stock.columns.tolist()}")

    ref_col = found_cols['ref']
    qty_col = found_cols['quantity']

    for col in stock.select_dtypes(include=['object']).columns:
        stock[col] = stock[col].astype('string').str.strip()

    stock[ref_col] = _normalize_ref(stock[ref_col])
    stock[qty_col] = pd.to_numeric(stock[qty_col], errors='coerce').fillna(0.0).round(2)

    return stock, ref_col, qty_col


def process_atelier(atelier_key: str, stock_file: dict, mov_file_path: str, month: str, overrides: dict | None = None) -> dict:
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}

    stock_file_path = stock_file.get('path') if isinstance(stock_file, dict) else None
    prev_stock_explicit_path = stock_file.get('prevPath') if isinstance(stock_file, dict) else None
    if not stock_file_path:
        return {'error': 'Missing stock file path', 'matches': [], 'discrepancies': []}

    overrides = overrides or {}
    sheet_override = overrides.get('sheetName')
    ref_override = overrides.get('refCol')
    qty_override = overrides.get('qtyCol')

    args = sheet_args[atelier_key]
    possible_sheets = args.get('sheet_name', [])
    if not isinstance(possible_sheets, list):
        possible_sheets = [possible_sheets]
    if sheet_override:
        possible_sheets = [sheet_override]

    try:
        # Current stock
        stock, stock_ref_col, stock_qty_col = load_stock(stock_file_path)

        # Previous stock (opening inventory)
        year = _infer_year_from_stock_filename(stock_file_path) or 2025
        prev_month, prev_year = prev_month_year(month, year)
        prev_stock_path = prev_stock_explicit_path or _find_neighbor_stock_file(stock_file_path, prev_month, prev_year)

        prev_stock_agg = pd.DataFrame(columns=['Ref', 'Prev_Stock_Qty'])
        if prev_stock_path and os.path.exists(prev_stock_path):
            prev_stock, prev_ref_col, prev_qty_col = load_stock(prev_stock_path)
            prev_stock_agg = prev_stock.groupby(prev_ref_col)[prev_qty_col].sum().reset_index()
            prev_stock_agg.rename(columns={prev_ref_col: 'Ref', prev_qty_col: 'Prev_Stock_Qty'}, inplace=True)
            prev_stock_agg['Prev_Stock_Qty'] = prev_stock_agg['Prev_Stock_Qty'].round(2)
        else:
            prev_stock_agg = pd.DataFrame({'Ref': [], 'Prev_Stock_Qty': []})

        # Movement
        mov = None
        used_sheet = None
        for sheet in possible_sheets:
            try:
                temp_df = pd.read_excel(mov_file_path, sheet_name=sheet, header=None)
                header_idx = _find_header_row(mov_file_path, sheet, mov_possible_col_names['date'])
                mov = pd.read_excel(mov_file_path, sheet_name=sheet, header=header_idx if header_idx else None)
                used_sheet = sheet
                break
            except ValueError:
                continue
            except Exception:
                continue

        if mov is None:
            return {'error': f"Could not read any of sheets {possible_sheets} in movement file", 'matches': [], 'discrepancies': []}

        found_cols = {}
        for col_type, possible_names in mov_possible_col_names.items():
            for name in possible_names:
                if name in mov.columns:
                    found_cols[col_type] = name
                    break

        ref_col = ref_override or found_cols.get('ref')
        qty_col = qty_override or found_cols.get('quantity')

        if not ref_col or not qty_col:
            return {'error': f"Could not find ref/quantity columns. Available: {mov.columns.tolist()}", 'matches': [], 'discrepancies': []}

        # Filter by Date: keep ONLY the target month for this unit.
        if 'date' in found_cols:
            date_col = found_cols['date']
            mov[date_col] = pd.to_datetime(mov[date_col], errors='coerce')
            target_month = int(month)
            target_year = year
            mask = (mov[date_col].dt.year == target_year) & (mov[date_col].dt.month == target_month)
            mov = mov[mask]

        for col in mov.select_dtypes(include=['object']).columns:
            mov[col] = mov[col].astype('string').str.strip()

        mov[ref_col] = _normalize_ref(mov[ref_col])
        mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce').fillna(0.0).round(2)

        mov_agg = mov.groupby(ref_col)[qty_col].sum().reset_index()
        mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
        mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)

        stock_agg = stock.groupby(stock_ref_col)[stock_qty_col].sum().reset_index()
        stock_agg.rename(columns={stock_ref_col: 'Ref', stock_qty_col: 'Stock_Qty'}, inplace=True)
        stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)

        expected_agg = pd.merge(prev_stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
        expected_agg['Expected_End_Qty'] = (expected_agg.get('Prev_Stock_Qty', 0) + expected_agg.get('Calc_Mov_Qty', 0)).round(2)

        comparison_df = pd.merge(stock_agg, expected_agg[['Ref', 'Expected_End_Qty']], on='Ref', how='outer').fillna(0)
        comparison_df.rename(columns={'Expected_End_Qty': 'Calc_Mov_Qty'}, inplace=True)
        comparison_df['Difference'] = (comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']).round(2)

        discrepancies = comparison_df[comparison_df['Difference'].abs() > 0.02].sort_values(by='Difference', ascending=False)
        matches = comparison_df[comparison_df['Difference'].abs() <= 0.02]

        return {'matches': matches.to_dict('records'), 'discrepancies': discrepancies.to_dict('records')}

    except Exception as e:
        return {'error': str(e), 'matches': [], 'discrepancies': []}


def process_all(stock_file, matched_files, month):
    results = {}
    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(atelier_key, stock_file, mov_file['path'], month, overrides=None)
    return results


def process_all_with_overrides(stock_file, matched_files, month, overrides):
    results = {}
    for atelier_key, mov_file in matched_files.items():
        atelier_overrides = overrides.get(atelier_key, {}) if overrides else {}
        results[atelier_key] = process_atelier(atelier_key, stock_file, mov_file['path'], month, overrides=atelier_overrides)
    return results
