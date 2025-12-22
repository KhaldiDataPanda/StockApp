"""
Oran Processor - Exact logic from test._oran2.py
This processor has:
- CSV support with delimiter detection
- Each atelier has its own stock sheet
"""

import pandas as pd
import os
import re
import csv


# Sheet arguments configuration
sheet_args = {
    'block': {'sheet_name': ['Movement Block'], 'stock_sheet': 'Stock Block'},
    'mousse': {'sheet_name': ['Movement Magasin mousse'], 'stock_sheet': 'STOCK MOUSSE'},
    'rouléss': {'sheet_name': ['Movement Roulés'], 'stock_sheet': 'STOCK ROULE'},
    'cardage': {'sheet_name': ['Atelier Cardage'], 'stock_sheet': 'STOCK FIBER'},
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'ref': ['Row Labels', 'Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'REF', 'REF PRODUIT', 'referance', 'البيان', 'المرجع', 'نوع', 'REFERNCE', 'التعيين'],
    'quantity': ['Q', 'Quantité', 'STOCK PV', 'STOCK FIBRE ', 'STOCK', 'STOCK ', 'STOCKS', 'ST-P', 'STOKS', 'STOCK U', 'المخزون', 'الكمية', 'العدد', 'STOCK/M', 'STOCK POIDS', 'باقي', 'Q-STOCKS', 'Q(U)', 'q(u)', 'Q-REEL', 'Somme de Q(U)', 'U'],
}

stock_qty_priority = ['q-reel', 'somme de q(u)', 'q(u)', 'Q', 'u', 'q-logicial', 'q-logiciale']


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def _norm_label(value):
    """Normalize to improve matching across casing, extra spaces, and embedded newlines."""
    return ' '.join(str(value).split()).strip().lower()


def _find_col(columns, candidates):
    """Find column with normalized matching"""
    col_map = {_norm_label(c): c for c in columns}
    for candidate in candidates:
        key = _norm_label(candidate)
        if key in col_map:
            return col_map[key]
    return None


def _read_csv_with_inferred_sep(path, header=None):
    """Try to detect delimiter and read CSV"""
    seps = [';', ',', '\t', '|']
    encodings = ['utf-8-sig', 'utf-8', 'cp1252']
    last_error = None

    def _sniff_delimiter(sample_text):
        try:
            dialect = csv.Sniffer().sniff(sample_text, delimiters=''.join(seps))
            return dialect.delimiter
        except Exception:
            return None

    for encoding in encodings:
        try:
            with open(path, 'rb') as f:
                sample_bytes = f.read(65536)
            sample_text = sample_bytes.decode(encoding, errors='replace')
        except Exception as e:
            last_error = e
            continue

        sniffed = _sniff_delimiter(sample_text)
        if sniffed:
            try:
                return pd.read_csv(path, sep=sniffed, header=header, encoding=encoding, engine='python')
            except Exception as e:
                last_error = e

        best_sep = None
        best_score = -1
        for sep in seps:
            try:
                preview = pd.read_csv(path, sep=sep, header=header, encoding=encoding, engine='python', nrows=50)
                score = int(preview.shape[1])
                if score > best_score:
                    best_score = score
                    best_sep = sep
            except Exception as e:
                last_error = e

        if best_sep is not None and best_score > 1:
            try:
                return pd.read_csv(path, sep=best_sep, header=header, encoding=encoding, engine='python')
            except Exception as e:
                last_error = e

        try:
            return pd.read_csv(path, sep=None, engine='python', header=header, encoding=encoding)
        except Exception as e:
            last_error = e

    raise last_error if last_error else Exception("Could not read CSV")


def _read_movement_file(path, possible_sheets):
    """Read movement file (CSV or Excel)"""
    ext = os.path.splitext(path)[1].lower()
    
    if ext == '.csv':
        temp_df = _read_csv_with_inferred_sep(path, header=None)
        header_idx = 0
        found_header = False

        for idx, row in temp_df.iterrows():
            row_values = [str(val) for val in row.values if pd.notna(val)]
            row_values_norm = {_norm_label(v) for v in row_values}
            if any(_norm_label(name) in row_values_norm for name in mov_possible_col_names['date']):
                header_idx = idx
                found_header = True
                break

        if not found_header:
            header_idx = temp_df.notna().sum(axis=1).idxmax()

        mov = _read_csv_with_inferred_sep(path, header=int(header_idx))
        return mov, 'CSV'

    # Excel fallback
    if isinstance(possible_sheets, str):
        possible_sheets = [possible_sheets]

    mov = None
    used_sheet = None
    for sheet in possible_sheets:
        try:
            temp_df = pd.read_excel(path, sheet_name=sheet, header=None)
            header_idx = 0
            found_header = False

            for idx, row in temp_df.iterrows():
                row_values = [str(val) for val in row.values if pd.notna(val)]
                row_values_norm = {_norm_label(v) for v in row_values}
                if any(_norm_label(name) in row_values_norm for name in mov_possible_col_names['date']):
                    header_idx = idx
                    found_header = True
                    break

            if found_header:
                mov = pd.read_excel(path, sheet_name=sheet, header=header_idx)
            else:
                header_idx = temp_df.notna().sum(axis=1).idxmax()
                mov = pd.read_excel(path, sheet_name=sheet, header=header_idx)

            used_sheet = sheet
            break
        except ValueError:
            continue
        except Exception:
            continue

    if mov is None:
        raise ValueError(f"Could not find any of the sheets {possible_sheets} in {path}")

    return mov, used_sheet


def process_atelier(atelier_key, stock_file_path, mov_file_path, month):
    """Process a single atelier and return matches/discrepancies"""
    
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}
    
    args = sheet_args[atelier_key]
    
    try:
        #----------------------------------------------------------
        # Processing Movement File
        #----------------------------------------------------------
        possible_sheets = args['sheet_name']
        mov, used_sheet = _read_movement_file(mov_file_path, possible_sheets)
        
        mov = mov.dropna(how="all")
        mov.columns = mov.columns.astype(str).str.strip()
        
        # Find columns (case/whitespace-insensitive)
        found_cols = {}
        for col_type, possible_names in mov_possible_col_names.items():
            matched = _find_col(mov.columns, possible_names)
            if matched is not None:
                found_cols[col_type] = matched
        
        if 'ref' not in found_cols:
            return {'error': f'Could not find ref column. Available: {mov.columns.tolist()}', 'matches': [], 'discrepancies': []}
        if 'quantity' not in found_cols:
            return {'error': f'Could not find quantity column. Available: {mov.columns.tolist()}', 'matches': [], 'discrepancies': []}
        
        mov_ref_col = found_cols['ref']
        mov_qty_col = found_cols['quantity']
        
        # Filter by Date if available
        if 'date' in found_cols:
            date_col = found_cols['date']
            mov[date_col] = pd.to_datetime(mov[date_col], errors='coerce')
            target_month = int(month)
            target_year = 2025
            
            mask = (mov[date_col].dt.year < target_year) | \
                   ((mov[date_col].dt.year == target_year) & (mov[date_col].dt.month <= target_month))
            mov = mov[mask]
        
        # Clean Movement Data
        object_columns = mov.select_dtypes(include=['object']).columns
        for col in object_columns:
            mov[col] = mov[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
        
        mov[mov_ref_col] = mov[mov_ref_col].astype('string')
        mov[mov_ref_col] = mov[mov_ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
        
        # Aggregate Movement
        mov[mov_qty_col] = pd.to_numeric(mov[mov_qty_col], errors='coerce').fillna(0.0)
        mov_agg = mov.groupby(mov_ref_col)[mov_qty_col].sum().reset_index()
        mov_agg.rename(columns={mov_ref_col: 'Ref', mov_qty_col: 'Calc_Mov_Qty'}, inplace=True)
        mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)
        
        #----------------------------------------------------------
        # Processing Stock File
        #----------------------------------------------------------
        stock_sheet_name = args['stock_sheet']
        
        stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=None)
        stock = stock.dropna(how="all")
        
        # Find header row with most non-empty values
        header_idx = stock.notna().sum(axis=1).idxmax()
        stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=header_idx)
        
        stock.columns = stock.columns.astype(str).str.strip()
        
        # Find reference column
        ref_col_candidates = [c for c in stock.columns if 'référence' in c.lower() or 'reference' in c.lower()]
        if not ref_col_candidates:
            ref_col_candidates = [c for c in stock.columns if 'ref' in c.lower()]
        
        stock_ref_col = ref_col_candidates[0] if ref_col_candidates else None
        
        if not stock_ref_col:
            return {'error': f'Could not find reference column in stock. Available: {stock.columns.tolist()}', 'matches': [], 'discrepancies': []}
        
        # Find quantity column with priority
        stock_qty_col = None
        for priority_name in stock_qty_priority:
            matched = _find_col(stock.columns, [priority_name])
            if matched:
                stock_qty_col = matched
                break
        
        if not stock_qty_col:
            qty_candidates = [c for c in stock.columns if any(q in c.lower() for q in ['quantité', 'quantite', 'qty', 'stock', 'q(u)'])]
            if qty_candidates:
                stock_qty_col = qty_candidates[0]
        
        if not stock_qty_col:
            return {'error': f'Could not find quantity column in stock. Available: {stock.columns.tolist()}', 'matches': [], 'discrepancies': []}
        
        # Clean stock
        stock[stock_ref_col] = stock[stock_ref_col].astype('string').str.strip()
        stock[stock_ref_col] = stock[stock_ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
        stock[stock_qty_col] = pd.to_numeric(stock[stock_qty_col], errors='coerce').fillna(0.0)
        
        # Aggregate Stock
        stock_agg = stock.groupby(stock_ref_col)[stock_qty_col].sum().reset_index()
        stock_agg.rename(columns={stock_ref_col: 'Ref', stock_qty_col: 'Stock_Qty'}, inplace=True)
        stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)
        
        # Comparison
        comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
        comparison_df['Difference'] = comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']
        
        # Results
        discrepancies = comparison_df[comparison_df['Difference'] != 0].sort_values(by='Difference', ascending=False)
        matches = comparison_df[comparison_df['Difference'] == 0]
        
        return {
            'matches': matches.to_dict('records'),
            'discrepancies': discrepancies.to_dict('records')
        }
        
    except Exception as e:
        return {'error': str(e), 'matches': [], 'discrepancies': []}


def process_all(stock_file, matched_files, month):
    """Process all matched files for Oran"""
    results = {}
    
    # Process each atelier
    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(atelier_key, stock_file['path'], mov_file['path'], month)
    
    return results
