"""
Fath5 Processor - Exact logic from test_fath5.py
"""

import pandas as pd
import os
import re


# Sheet arguments configuration
sheet_args = {
    'bonda': {'sheet_name': ['ATELLIER COUATE'], 'localisation': ['ATELIER BONDA 3D']},
    'orillier': {'sheet_name': ['MOV'], 'localisation': ["ATELIER CONFECTION D'ORILLIER"]},
    'confiction': {'sheet_name': ['ATT CONFECTION'], 'localisation': ['ATELIER CONFICTION']},
    'couette fini': {'sheet_name': ['ATELLIER COUATE'], 'localisation': ['ATTELLIER COUETTE  FINI']},
    'semi fini': {'sheet_name': ['ATELLIER COUATE'], 'localisation': ['ATELIER COUETTE SEMI FINI']},
    'rouli': {'sheet_name': ['MOV'], 'localisation': ['ATELIER MATELAS ROULI ']},
    'block': {'sheet_name': ['MOUV'], 'localisation': ['MAGASIN DE BLOCK']},
    'gratage': {'sheet_name': ['MOUV'], 'localisation': ['ATELIER GRATTAGE']},
    'coupage': {'sheet_name': ['MOUV'], 'localisation': ['ATELIER COUPAGE']},
    'comersial': {'sheet_name': ['MOV'], 'localisation': ['MAGASIN COMMERCIAL']},
    'secondaire': {'sheet_name': ['MAG'], 'localisation': ['MAGASIN SECONDAIRE EL FATEH 05'], 'mov cols': [-1, 'Q-STOCKS']},
    'ouate': {'sheet_name': ['MOUV'], 'localisation': ['ATELIER OUATE']},
    'outin': {'sheet_name': ['MOUVEMENT'], 'localisation': ['ATELIER OUATENAGE']},
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'التاريخ'],
    'ref': ['Row Labels', 'Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'REF', 'REF PRODUIT', 'referance', 'البيان', 'المرجع', 'نوع', 'REFERNCE', 'التعيين'],
    'quantity': ['Quantité', 'STOCK PV', 'STOCK FIBRE ', 'STOCK', 'STOCK ', 'STOCKS', 'ST-P', 'STOKS', 'STOCK U', 'المخزون', 'الكمية', 'العدد', 'STOCK/M', 'STOCK POIDS', 'باقي', 'Q-STOCKS'],
}


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def load_stock(stock_file_path, stock_sheet_name=None):
    """Load and prepare stock data from Fath5"""
    # Find appropriate sheet
    try:
        xl = pd.ExcelFile(stock_file_path)
        available_sheets = xl.sheet_names
        
        # Try MOUV first, then STOCK, then first sheet
        sheet_to_use = None
        for preferred in ['MOUV', 'MOV', 'STOCK']:
            for sheet in available_sheets:
                if preferred.upper() in sheet.upper():
                    sheet_to_use = sheet
                    break
            if sheet_to_use:
                break
        
        if sheet_to_use is None:
            sheet_to_use = available_sheets[0]
    except Exception as e:
        raise ValueError(f"Could not read stock file: {stock_file_path} - {str(e)}")
    
    # Scan for header row containing a date column
    temp_stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=None)
    header_idx = 0
    found_header = False
    
    for idx, row in temp_stock.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        if any(name in row_values for name in mov_possible_col_names['date']):
            header_idx = idx
            found_header = True
            break
    
    if found_header:
        stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=header_idx)
    else:
        stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=0)
    
    stock = stock.dropna(how="all")
    
    # Ensure types for Quantity and Date
    if 'QUANTITE' in stock.columns:
        stock['QUANTITE'] = pd.to_numeric(stock['QUANTITE'], errors='coerce').fillna(0.0)
        stock['QUANTITE'] = stock['QUANTITE'].round(2)
    
    date_cols = [col for col in stock.columns if 'date' in col.lower()]
    for col in date_cols:
        stock[col] = pd.to_datetime(stock[col], errors='coerce')
    
    object_columns = stock.select_dtypes(include=['object']).columns
    for col in object_columns:
        stock[col] = stock[col].astype(str).str.strip()
    
    # Find REF column - Fath5 uses 'REF' not 'REFERENCE'
    ref_col_name = None
    for col in stock.columns:
        if 'REF' in str(col).upper():
            ref_col_name = col
            break
    
    if ref_col_name:
        stock[ref_col_name] = stock[ref_col_name].astype('string')
        stock[ref_col_name] = stock[ref_col_name].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
        # Rename to standard REFERENCE for consistency
        stock.rename(columns={ref_col_name: 'REFERENCE'}, inplace=True)
    
    return stock


def process_atelier(atelier_key, stock_df, mov_file_path, month):
    """Process a single atelier and return matches/discrepancies"""
    
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}
    
    args = sheet_args[atelier_key]
    
    try:
        # Handle multiple possible sheet names
        possible_sheets = args['sheet_name']
        if isinstance(possible_sheets, str):
            possible_sheets = [possible_sheets]
        
        mov = None
        used_sheet = None
        
        for sheet in possible_sheets:
            try:
                # Read Movement File - scan for header row
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
            except Exception as e:
                continue
        
        if mov is None:
            return {
                'error': f"Could not find any of the sheets {possible_sheets}",
                'matches': [],
                'discrepancies': []
            }
        
        # Find Columns with optional override
        found_cols = {}
        for col_type, possible_names in mov_possible_col_names.items():
            for name in possible_names:
                if name in mov.columns:
                    found_cols[col_type] = name
                    break
        
        # Handle column override from args
        mov_cols_override = args.get('mov cols')
        override_ref = None
        override_qty = None
        
        if isinstance(mov_cols_override, (list, tuple)) and len(mov_cols_override) == 2:
            override_ref, override_qty = mov_cols_override
            if override_ref == -1:
                override_ref = None
            if override_qty == -1:
                override_qty = None
        
        # Determine final ref/qty column names
        if override_ref is not None:
            ref_col = override_ref
        elif 'ref' in found_cols:
            ref_col = found_cols['ref']
        else:
            return {
                'error': f"Could not find ref column. Available: {mov.columns.tolist()}",
                'matches': [],
                'discrepancies': []
            }
        
        if override_qty is not None:
            qty_col = override_qty
        elif 'quantity' in found_cols:
            qty_col = found_cols['quantity']
        else:
            return {
                'error': f"Could not find quantity column. Available: {mov.columns.tolist()}",
                'matches': [],
                'discrepancies': []
            }
        
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
        object_columns_mov = mov.select_dtypes(include=['object']).columns
        for col in object_columns_mov:
            mov[col] = mov[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
        
        mov[ref_col] = mov[ref_col].astype('string')
        mov[ref_col] = mov[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
        
        # Group Movement
        mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce').fillna(0.0)
        mov[qty_col] = mov[qty_col].round(2)
        mov_agg = mov.groupby(ref_col)[qty_col].sum().reset_index()
        mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
        mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)
        
        # Filter Stock by localisation
        localisations = args['localisation']
        stock_filtered = stock_df[stock_df['LOCALISATION'].isin(localisations)].copy()
        
        # Group Stock
        qty_stock_col = 'QUANTITE' if 'QUANTITE' in stock_df.columns else stock_df.columns[stock_df.columns.str.contains('QUANT', case=False)][0] if any(stock_df.columns.str.contains('QUANT', case=False)) else None
        
        if qty_stock_col is None:
            return {'error': 'Could not find quantity column in stock', 'matches': [], 'discrepancies': []}
        
        stock_agg = stock_filtered.groupby('REFERENCE')[qty_stock_col].sum().reset_index()
        stock_agg.rename(columns={'REFERENCE': 'Ref', qty_stock_col: 'Stock_Qty'}, inplace=True)
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
    """Process all matched files for Fath5"""
    results = {}
    
    # Load stock once
    try:
        stock_df = load_stock(stock_file['path'])
    except Exception as e:
        return {'_error': f'Failed to load stock file: {str(e)}'}
    
    # Process each atelier
    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(atelier_key, stock_df, mov_file['path'], month)
    
    return results
