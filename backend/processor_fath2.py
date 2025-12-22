"""
Fath2 Processor - Exact logic from test_fath2.py
"""

import pandas as pd
import os
import re


# Sheet arguments configuration
sheet_args = {
    'ouate 01': {'sheet_name': ['الحركة اليومية'], 'localisation': ['ATT OUATE 01']},
    'ouate 02': {'sheet_name': ['الحركة اليومية'], 'localisation': ['ATT OUATE 02']},
    'sfifa': {'sheet_name': ['MOUVEMENT 2024'], 'localisation': ['ATT SFIFA', 'MGZ TRANSFERT']},
    'mgz plasic': {'sheet_name': ['MOVEMENT MAGASAIN PLASTIQUE'], 'localisation': ['MGZ PLASTIG']},
    'secondaire': {'sheet_name': ['MOUV'], 'localisation': ['MAG UNITE']},
    '-dechet': {'sheet_name': ['MOUVEMENT'], 'localisation': ['MAG DECHET']},
    'commercial': {'sheet_name': ['MOUVEMENT'], 'localisation': ['SERVICE COMMERCIAL']},
    'فيبر': {'sheet_name': ['حركة الفيبر اليومية'], 'localisation': [' ateliers FIBRE cardi']},
    'plastique': {'sheet_name': ['MOVEMONTE DE ATT PLASTIQUE', 'MOVEMONTE'], 'localisation': ['ATELIER PLASTIQUE']},
    'tissu': {'sheet_name': ['MOUVEMENT 2024'], 'localisation': ['ATT- TISSU']},
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date'],
    'ref': ['Row Labels', 'Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'REF', 'REF PRODUIT'],
    'quantity': ['Quantité', 'STOCK PV', 'STOCK', 'STOCKS', 'ST-P', 'STOKS'],
}


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def load_stock(stock_file_path, stock_sheet_name=None):
    """Load and prepare stock data from Fath2"""
    # Find appropriate sheet
    try:
        xl = pd.ExcelFile(stock_file_path)
        available_sheets = xl.sheet_names
        
        # Try MOV first, then STOCK, then first sheet
        sheet_to_use = None
        for preferred in ['MOV', 'STOCK']:
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
    
    if 'REFERENCE' in stock.columns:
        stock['REFERENCE'] = stock['REFERENCE'].astype('string')
        stock['REFERENCE'] = stock['REFERENCE'].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
    
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
        
        # Find Columns
        found_cols = {}
        for col_type, possible_names in mov_possible_col_names.items():
            for name in possible_names:
                if name in mov.columns:
                    found_cols[col_type] = name
                    break
        
        if 'ref' not in found_cols or 'quantity' not in found_cols:
            return {
                'error': f"Could not find ref or quantity columns. Available: {mov.columns.tolist()}",
                'matches': [],
                'discrepancies': []
            }
        
        ref_col = found_cols['ref']
        qty_col = found_cols['quantity']
        
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
        stock_agg = stock_filtered.groupby('REFERENCE')['QUANTITE'].sum().reset_index()
        stock_agg.rename(columns={'REFERENCE': 'Ref', 'QUANTITE': 'Stock_Qty'}, inplace=True)
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
    """Process all matched files for Fath2"""
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
