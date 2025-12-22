"""
Fath1 Processor - Exact logic from test_fath1.py
"""

import pandas as pd
import os
import re


# Sheet arguments configuration
sheet_args = {
    'femme 01': {'sheet_name': 'الحركة اليومية', 'localisation': ['ATELIER COUTURE FEMME 01']},
    'femme 02': {'sheet_name': 'الحركة اليومية', 'localisation': ['ATELIER COUTURE FEMME 02']},
    'bourde': {'sheet_name': 'MOVEMENT BORDI', 'localisation': ['ATT COUPAGE BOURDEE']},
    'coupage': {'sheet_name': 'MOVEMENT', 'localisation': ['ATELIER DE COUPAGE A', 'ATELIER DE COUPAGE B', 'ATT COUPAGE ROULER', 'ATELIER DE ROULOUX & ACCOUDOIRE']},
    '+croute': {'sheet_name': 'MOUVMENT', 'localisation': ['ATELIER DE DECHETS']},
    'grattage+': {'sheet_name': 'MOUVEMENT', 'localisation': ['ATELIER DE DECHETS']},
    'produt': {'sheet_name': 'MOVEMMENT', 'localisation': ['ATT- PRODUIT EL FATH 01']},
    '-grattage': {'sheet_name': 'ATT GRATTAGE', 'localisation': ['ATTELIER GRATTAGE A', 'ATTELIER GRATTAGE B', 'ATT GRATTAGE C', 'ATT GRATTAGE BORDER', 'ATT GRATTAGE ROULER']},
    'rouler': {'sheet_name': 'MOVEMENT ROLI', 'localisation': ['ATT  ROULER']},
    'conftection rouli': {'sheet_name': 'MOUVEMENT ATELIER CONFECTION', 'localisation': ['ATTELIER CONFECTION ROULI']},
    'brodri': {'sheet_name': 'MOUVEMENT', 'localisation': ['ATTELIER BRODRIE', 'ATTELIER BRODERI']},
    'conftection bourdi': {'sheet_name': 'MOUVEMENT ATELIER CONFECTION', 'localisation': ['ATELIER CONFECTION BOURDI']},
    'fiber cardi': {'sheet_name': 'الحركة اليومية', 'localisation': ['ATTELLIER  FIBER CARDI']},
    'orielle': {'sheet_name': 'الحركة اليومية', 'localisation': ["ATTELIER D'ORIELLER"]},
    'bloc': {'sheet_name': 'MOUVEM 09', 'localisation': ['MAGASINE DE BLOCS']},
    'magaza bourdi': {'sheet_name': 'مخزن البوردي (بلاستيك+ساكوشة)', 'localisation': ['MAGASIN BOURDI']},
    'secondaire': {'sheet_name': 'MAGASIN', 'localisation': ['MAGASIN SECONDAIRE(FATH1)']},
    'mov-com': {'sheet_name': 'MOV', 'localisation': ['MAGASIN COMMERCIAL']}
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date'],
    'ref': ['Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'Référence\n'],
    'quantity': ['STOCK PV', 'STOCK', 'STOCKS', 'ST-P', 'ST-PV'],
}


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def load_stock(stock_file_path, stock_sheet_name=None):
    """Load and prepare stock data from Fath1"""
    # Try to read Excel and find appropriate sheet
    try:
        xl = pd.ExcelFile(stock_file_path)
        available_sheets = xl.sheet_names
        
        # Try to find a sheet with STOCK in the name
        sheet_to_use = None
        for sheet in available_sheets:
            if 'STOCK' in sheet.upper():
                sheet_to_use = sheet
                break
        
        # If no STOCK sheet found, use first sheet
        if sheet_to_use is None:
            sheet_to_use = available_sheets[0]
        
        stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=None)
    except Exception as e:
        raise ValueError(f"Could not read stock file: {stock_file_path} - {str(e)}")
    
    stock = stock.dropna(how="all")
    
    # Find the row with the MOST non-empty values (this should be your header row)
    header_idx = stock.notna().sum(axis=1).idxmax()
    stock = pd.read_excel(stock_file_path, sheet_name=sheet_to_use, header=header_idx)
    
    # Clean object columns
    object_columns = stock.select_dtypes(include=['object']).columns
    for col in object_columns:
        if col in stock.columns:
            stock[col] = stock[col].astype(str).str.strip()
    
    # Normalize reference
    if 'REFERENCE' in stock.columns:
        stock['REFERENCE'] = stock['REFERENCE'].astype('string')
        stock['REFERENCE'] = stock['REFERENCE'].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
    
    # Convert quantity
    if 'QUANTITE' in stock.columns:
        stock['QUANTITE'] = pd.to_numeric(stock['QUANTITE'], errors='coerce')
    
    return stock


def process_atelier(atelier_key, stock_df, mov_file_path, month):
    """Process a single atelier and return matches/discrepancies"""
    
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}
    
    args = sheet_args[atelier_key]
    
    try:
        # Read Movement File - scan for header row containing a date column
        temp_df = pd.read_excel(mov_file_path, sheet_name=args['sheet_name'], header=None)
        header_idx = 0
        found_header = False
        
        for idx, row in temp_df.iterrows():
            row_values = [str(val).strip() for val in row.values if pd.notna(val)]
            if any(name in row_values for name in mov_possible_col_names['date']):
                header_idx = idx
                found_header = True
                break
        
        if found_header:
            mov = pd.read_excel(mov_file_path, sheet_name=args['sheet_name'], header=header_idx)
        else:
            mov = pd.read_excel(mov_file_path, sheet_name=args['sheet_name'])
        
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
        
        # Clean Movement Data
        object_columns_mov = mov.select_dtypes(include=['object']).columns
        for col in object_columns_mov:
            mov[col] = mov[col].astype(str).str.strip()
        
        mov[ref_col] = mov[ref_col].astype('string')
        mov[ref_col] = mov[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
        mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce')
        
        # Group Movement
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
    """Process all matched files for Fath1"""
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
