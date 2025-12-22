"""
Fibre Processor - Exact logic from fibre.py
This processor has:
- Dual-column filtering (LOCAL + LOCALISATION)
- Multi-job support for nested sheet configurations
- Tuple-based localisations
"""

import pandas as pd
import os
import re


# Sheet arguments configuration
sheet_args = {
    'drafter': {'sheet_name': ['Drafter'], 'localisation': [('At-Fibre2', 'DRAFTER')]},
    'extredeuse': {'sheet_name': ['ATELLIER COUATE'], 'localisation': [('At-Fibre2', 'EXTREDEUSE')]},
    'filiére': {'sheet_name': ['ATELLIER COUATE'], 'localisation': [('At-Fibre2', 'Filiére')]},
    'carding': {'sheet_name': ['الحركةاليومية'], 'localisation': ['AT-CARDING'], 'mov cols': ['Référence', 'Quantité']},
    'magaisain pet': {'sheet_name': ['Mouvement'], 'localisation': ['MAGASIN-PET'], 'mov cols': ['Référence', 'Quantité']},
    'magaisain fibre': {'sheet_name': [['MOUVEMENT'], ['UNITE ETPH']], 'localisation': [['MAGASIN'], ['MAGASIN-LAROBI']], 'mov cols': [['REF PRODUIT', 'Quantité'], ['REF PRODUIT', 'S REEL']]},
    'magaisain commercial': {'sheet_name': ['Mouvement'], 'localisation': ['MAGASIN-Commerciale'], 'mov cols': ['Référence', 'Quantity']},
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'ref': ['Row Labels', 'Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'REF', 'REF PRODUIT', 'referance', 'البيان', 'المرجع', 'نوع', 'REFERNCE', 'التعيين'],
    'quantity': ['Quantité', 'STOCK PV', 'STOCK FIBRE ', 'STOCK', 'STOCK ', 'STOCKS', 'ST-P', 'STOKS', 'STOCK U', 'المخزون', 'الكمية', 'العدد', 'STOCK/M', 'STOCK POIDS', 'باقي', 'Q-STOCKS'],
}


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def _norm_col_name(value):
    if value is None:
        return ""
    return str(value).strip().casefold()


def find_column(df, candidates):
    """Find an existing column in df whose normalized name matches any candidate."""
    if df is None or df.empty:
        return None

    normalized_to_actual = {_norm_col_name(c): c for c in df.columns}
    for candidate in candidates:
        key = _norm_col_name(candidate)
        if key in normalized_to_actual:
            return normalized_to_actual[key]
    return None


def normalize_to_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


def is_tuple_localisations(localisations):
    return any(isinstance(item, tuple) and len(item) == 2 for item in localisations)


def filter_stock_by_localisation(stock_df, localisations_spec, local_col, localisation_col):
    """Filter stock by localisation (supports tuple-based filtering)"""
    localisations = normalize_to_list(localisations_spec)
    if not localisations:
        return stock_df.copy()

    if is_tuple_localisations(localisations):
        if not localisation_col:
            raise KeyError("LOCALISATION column required for tuple-based localisation filtering")
        mask = pd.Series(False, index=stock_df.index)
        for local_value, localisation_value in localisations:
            mask = mask | (
                (stock_df[local_col].astype("string") == str(local_value))
                & (stock_df[localisation_col].astype("string") == str(localisation_value))
            )
        return stock_df[mask].copy()

    # Default: single-column LOCAL filtering
    localisations_str = [str(x) for x in localisations]
    return stock_df[stock_df[local_col].isin(localisations_str)].copy()


def build_jobs_from_args(args):
    """Normalize a sheet_args entry into one or more movement jobs."""
    sheets_spec = args.get("sheet_name")
    loc_spec = args.get("localisation")
    mov_cols_spec = args.get("mov cols")

    multi = False
    if isinstance(sheets_spec, list) and any(isinstance(x, (list, tuple)) for x in sheets_spec):
        multi = True
    if isinstance(loc_spec, list) and any(isinstance(x, (list, tuple)) for x in loc_spec):
        if not (isinstance(loc_spec, list) and loc_spec and all(isinstance(x, tuple) for x in loc_spec)):
            multi = True
    if isinstance(mov_cols_spec, list) and mov_cols_spec and any(isinstance(x, (list, tuple)) for x in mov_cols_spec):
        if len(mov_cols_spec) > 0 and isinstance(mov_cols_spec[0], (list, tuple)) and len(mov_cols_spec) > 1:
            multi = True

    jobs = []

    if not multi:
        possible_sheets = normalize_to_list(sheets_spec)
        jobs.append({
            "possible_sheets": [s for s in possible_sheets if s is not None],
            "localisation": loc_spec,
            "mov_cols": mov_cols_spec,
        })
        return jobs

    # Multi job
    sheet_groups = normalize_to_list(sheets_spec)
    loc_groups = normalize_to_list(loc_spec)
    mov_cols_groups = normalize_to_list(mov_cols_spec)

    if len(loc_groups) != len(sheet_groups):
        raise ValueError(f"Multi-sheet args require aligned lengths")

    for i in range(len(sheet_groups)):
        possible_sheets = normalize_to_list(sheet_groups[i])
        localisation = loc_groups[i]
        mov_cols = None if mov_cols_spec is None else mov_cols_groups[i] if i < len(mov_cols_groups) else None

        jobs.append({
            "possible_sheets": [s for s in possible_sheets if s is not None],
            "localisation": localisation,
            "mov_cols": mov_cols,
        })

    return jobs


def load_stock(stock_file_path, stock_sheet_name='STOCKS GLOBALE'):
    """Load and prepare stock data from Fibre"""
    # Scan for header row containing a date column
    temp_stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=None)
    header_idx = 0
    found_header = False

    for idx, row in temp_stock.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        if any(name in row_values for name in mov_possible_col_names['date']):
            header_idx = idx
            found_header = True
            break

    if found_header:
        stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=header_idx)
    else:
        stock = pd.read_excel(stock_file_path, sheet_name=stock_sheet_name, header=0)

    stock = stock.dropna(how="all")

    # Find columns
    stock_local_col = find_column(stock, ["LOCAL", "Local", "LOCAL "])
    stock_localisation_col = find_column(stock, ["LOCALISATION", "LOCALISATION ", "Localisation"])
    stock_ref_col = find_column(stock, ["PRODUIT ", " PRODUIT ", "PRODUIT", "REF", "REF "])
    stock_qty_col = find_column(stock, ["S REEL", "S REEL ", "S RÉEL", "QUANTITE", "QUANTITÉ", "QUANTITE "])

    if not stock_local_col:
        raise KeyError(f"Stock file is missing LOCAL column. Available: {stock.columns.tolist()}")
    if not stock_ref_col:
        raise KeyError(f"Stock file is missing reference column. Available: {stock.columns.tolist()}")
    if not stock_qty_col:
        raise KeyError(f"Stock file is missing quantity column. Available: {stock.columns.tolist()}")

    # Ensure types
    stock[stock_qty_col] = pd.to_numeric(stock[stock_qty_col], errors='coerce').fillna(0.0)
    stock[stock_qty_col] = stock[stock_qty_col].round(2)

    object_columns = stock.select_dtypes(include=['object']).columns
    for col in object_columns:
        stock[col] = stock[col].astype('string').str.strip()

    stock[stock_ref_col] = stock[stock_ref_col].astype('string')
    stock[stock_ref_col] = stock[stock_ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)

    return stock, stock_local_col, stock_localisation_col, stock_ref_col, stock_qty_col


def process_atelier(atelier_key, stock_df, stock_local_col, stock_localisation_col, stock_ref_col, stock_qty_col, mov_file_path, month):
    """Process a single atelier and return matches/discrepancies"""

    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}

    args = sheet_args[atelier_key]

    try:
        jobs = build_jobs_from_args(args)
    except Exception as e:
        return {'error': f'Error building jobs: {str(e)}', 'matches': [], 'discrepancies': []}

    all_matches = []
    all_discrepancies = []

    for job_idx, job in enumerate(jobs):
        possible_sheets = job["possible_sheets"]
        localisations_spec = job["localisation"]
        mov_cols_override = job.get("mov_cols")

        try:
            mov = None
            used_sheet = None

            for sheet in possible_sheets:
                try:
                    temp_df = pd.read_excel(mov_file_path, sheet_name=sheet, header=None)
                    header_idx = 0
                    found_header = False

                    for idx, row in temp_df.iterrows():
                        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
                        if any(name in row_values for name in mov_possible_col_names['date']):
                            header_idx = idx
                            found_header = True
                            break

                    if found_header:
                        mov = pd.read_excel(mov_file_path, sheet_name=sheet, header=header_idx)
                    else:
                        mov = pd.read_excel(mov_file_path, sheet_name=sheet)

                    used_sheet = sheet
                    break
                except ValueError:
                    continue
                except Exception:
                    continue

            if mov is None:
                continue

            # Find Columns
            found_cols = {}
            mov_cols_norm = {_norm_col_name(c): c for c in mov.columns}
            for col_type, possible_names in mov_possible_col_names.items():
                for name in possible_names:
                    candidate_norm = _norm_col_name(name)
                    if candidate_norm in mov_cols_norm:
                        found_cols[col_type] = mov_cols_norm[candidate_norm]
                        break

            override_ref = None
            override_qty = None

            if isinstance(mov_cols_override, (list, tuple)) and len(mov_cols_override) == 2:
                override_ref, override_qty = mov_cols_override

                if override_ref == -1:
                    override_ref = None
                if override_qty == -1:
                    override_qty = None

            # Determine final column names
            if override_ref is not None and _norm_col_name(override_ref) in mov_cols_norm:
                ref_col = mov_cols_norm[_norm_col_name(override_ref)]
            elif 'ref' in found_cols:
                ref_col = found_cols['ref']
            else:
                continue

            if override_qty is not None and _norm_col_name(override_qty) in mov_cols_norm:
                qty_col = mov_cols_norm[_norm_col_name(override_qty)]
            elif 'quantity' in found_cols:
                qty_col = found_cols['quantity']
            else:
                continue

            # Filter by Date
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

            mov[ref_col] = mov[ref_col].astype('string')
            mov[ref_col] = mov[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
            mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce').fillna(0.0)
            mov[qty_col] = mov[qty_col].round(2)

            # Aggregate Movement
            mov_agg = mov.groupby(ref_col)[qty_col].sum().reset_index()
            mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
            mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)

            # Filter Stock by localisation
            stock_filtered = filter_stock_by_localisation(stock_df, localisations_spec, stock_local_col, stock_localisation_col)

            # Aggregate Stock
            stock_agg = stock_filtered.groupby(stock_ref_col)[stock_qty_col].sum().reset_index()
            stock_agg.rename(columns={stock_ref_col: 'Ref', stock_qty_col: 'Stock_Qty'}, inplace=True)
            stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)

            # Comparison
            comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
            comparison_df['Difference'] = comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']

            discrepancies = comparison_df[comparison_df['Difference'] != 0]
            matches = comparison_df[comparison_df['Difference'] == 0]

            all_matches.extend(matches.to_dict('records'))
            all_discrepancies.extend(discrepancies.to_dict('records'))

        except Exception as e:
            continue

    return {
        'matches': all_matches,
        'discrepancies': sorted(all_discrepancies, key=lambda x: abs(x.get('Difference', 0)), reverse=True)
    }


def process_all(stock_file, matched_files, month):
    """Process all matched files for Fibre"""
    results = {}

    # Load stock once
    try:
        stock_df, stock_local_col, stock_localisation_col, stock_ref_col, stock_qty_col = load_stock(stock_file['path'])
    except Exception as e:
        return {'_error': f'Failed to load stock file: {str(e)}'}

    # Process each atelier
    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(
            atelier_key, stock_df, stock_local_col, stock_localisation_col,
            stock_ref_col, stock_qty_col, mov_file['path'], month
        )

    return results
