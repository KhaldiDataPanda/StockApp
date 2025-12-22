"""
Larbaa Processor - Exact logic from test_larba.py
This is the most complex processor with:
- Sheet-based stock (not localisation-based)
- Paired mode for multiple sheet groups
- US/EU date format detection
"""

import pandas as pd
import os
import re


# Sheet arguments configuration
sheet_args = {
    'atelier découpage': {'sheet_name': [' ATT-DECOUPAGE'], 'stock_sheet': {'sheets': ['ATT DECOUPAGE']}},
    'coupage ': {'sheet_name': ['ATT COUPAGE', 'ATT-COUP'], 'stock_sheet': {'sheets': ['ATT COUPAGE']}},
    'roulés.entré': {'sheet_name': ['Entré Mousse'], 'stock_sheet': {'sheets': ['ATT MATELAS ROULEE 01']}},
    'roulés.sortie': {'sheet_name': [['Sortie Atelier 01', 'Matière consommable']], 'stock_sheet': {'sheets': [['ATT MATELAS ROULEE 01']]}},
    'oreiller': {'sheet_name': ['OR'], 'stock_sheet': {'sheets': ['ATT ORIELE', 'MAG COUETTE+ORIELE']}},
    'couette': {'sheet_name': ['co', 'COUETTE'], 'stock_sheet': {'sheets': ['MAG COUETTE+ORIELE']}},
    'magasin blocs': {'sheet_name': ['Mouvements', 'Mouvement'], 'stock_sheet': {'sheets': ['STOCK BLOCK']}},
    'magasin fibre': {'sheet_name': ['Mouvements'], 'stock_sheet': {'sheets': ['STOCK FIBRE']}},
    'magasin ouate': {'sheet_name': ['Magasin Ouate'], 'stock_sheet': {'sheets': ['MAG OUATE']}},
    'magasin mousse': {'sheet_name': [' MOUVEMENT01'], 'stock_sheet': {'sheets': ['MAG MOUSSE']}},
    'magasin roules': {'sheet_name': ['Roulés'], 'stock_sheet': {'sheets': ['MAG ROULEE']}},
    'grattage': {'sheet_name': ['MAG DE GRATAG'], 'stock_sheet': {'sheets': ['ATT-GRATTAGE']}},
    'accessoire': {'sheet_name': ['Mouvements'], 'stock_sheet': {'sheets': ['MAGASN CENTRAL']}},
    'piec': {'sheet_name': ['Sheet1'], 'stock_sheet': {'sheets': ['PIECE']}},
    '(ouate': {'sheet_name': [['Opérations Ouate'], ['Opération Cardinage'], ['Opérations Ouatinage']], 
               'stock_sheet': {'sheets': [['ATT-PROD OUATE'], ['ATT-FIBRE CARDEE'], ['ATT-OUATINAGE']], 
                              'quantity': [['QUANTITE/KG'], ['QUANTITE KG'], ['QUANTITE/M']]}},
    'sortie mousse couture': {'sheet_name': ['Sortie Mouse Couture', "Sortie Mousse Couture"], 'stock_sheet': {'sheets': ['ATT COUPAGE COUTURE']}},
    ' pet ': {'sheet_name': ['MOUVMENT', 'MOUVEMENT'], 'stock_sheet': {'sheets': ['PET']}},
}

mov_possible_col_names = {
    'date': ['Date', 'DATE', 'date', 'LA DATE'],
    'ref': ['Row Labels', 'Référence (Bir Khadem)', 'Référence\nFournisseur', 'REFERENCE', 'RÉFÉRENCE', 'reference', 'Référence', 'REF', 'Ref', 'REF PRODUIT', 'REFERANCE', 'PRODUIT', 'REFFERENCE', 'réfferance'],
    'quantity': ['SOMME', 'Quantité', 'STOCK PV', 'STOCK', 'STOCKS', 'ST-P', 'STOKS', 'Stock', 'Stock(Kg)', ' Quantité', 'STOCK KG', 'STOCKS KG', 'STOCKS M', 'stock'],
}

stock_possible_col_names = {
    'date': ['Date', 'DATE', 'date'],
    'ref': ['REFERENCE', 'Référence', 'REF', 'reference'],
    'quantity': ['QUANTITE/KG', 'QUANTITE KG', 'QUANTITE', 'Quantité', 'QTE', 'STOCK'],
}


def get_ateliers():
    """Return list of atelier keywords"""
    return list(sheet_args.keys())


def detect_date_format_us_vs_eu(date_series, sample_size=200):
    """Detect whether a date series is more likely US (MM/DD/YYYY) or EU (DD/MM/YYYY)."""
    if date_series is None:
        return 'unknown'

    try:
        if pd.api.types.is_datetime64_any_dtype(date_series):
            return 'unknown'
    except Exception:
        pass

    s = pd.Series(date_series).dropna()
    if s.empty:
        return 'unknown'

    s = s.head(sample_size).astype('string').str.strip()
    s = s[s != '']
    if s.empty:
        return 'unknown'

    md_re = re.compile(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})')
    iso_re = re.compile(r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}')
    us_signals = 0
    eu_signals = 0
    iso_signals = 0

    for val in s.tolist():
        if not isinstance(val, str):
            continue
        val = val.strip()
        if not val:
            continue

        if iso_re.match(val):
            iso_signals += 1
            continue

        m = md_re.match(val)
        if not m:
            continue
        a = int(m.group(1))
        b = int(m.group(2))

        if a > 12 and b <= 12:
            eu_signals += 1
        elif b > 12 and a <= 12:
            us_signals += 1

    if us_signals > eu_signals and us_signals > 0:
        return 'us'
    if eu_signals > us_signals and eu_signals > 0:
        return 'eu'

    if iso_signals > 0 and iso_signals >= (len(s) * 0.5):
        return 'iso'

    return 'eu'


def parse_dates_normalized_eu(df, date_col):
    """Parse a date column, detecting US vs EU string formats."""
    fmt = detect_date_format_us_vs_eu(df[date_col])
    if fmt == 'us':
        parsed = pd.to_datetime(df[date_col], errors='coerce', dayfirst=False)
    elif fmt == 'iso':
        parsed = pd.to_datetime(df[date_col], errors='coerce', dayfirst=False)
    else:
        parsed = pd.to_datetime(df[date_col], errors='coerce', dayfirst=True)

    df[date_col] = parsed
    return parsed


def find_header_row(temp_df, possible_col_names):
    """Find the header row in a dataframe."""
    header_idx = 0
    found_header = False
    
    date_patterns = [
        r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',
        r'\d{4}[/-]\d{1,2}[/-]\d{1,2}',
    ]
    
    for idx, row in temp_df.head(50).iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        for val in row_values:
            if 'date' in val.lower():
                header_idx = idx
                found_header = True
                break
        if found_header:
            break
        if any(name in row_values for name in possible_col_names['date']):
            header_idx = idx
            found_header = True
            break
    
    if found_header:
        return header_idx, found_header
    
    for idx, row in temp_df.head(50).iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        has_ref = any(name in row_values for name in possible_col_names['ref'])
        has_qty = any(name in row_values for name in possible_col_names['quantity'])
        if has_ref or has_qty:
            header_idx = idx
            found_header = True
            return header_idx, found_header
    
    return header_idx, found_header


def read_stock_from_sheets(stock_file, sheet_names, quantity_cols=None):
    """Read and combine stock data from multiple sheets in the stock file."""
    all_stock_data = []
    
    for idx, sheet_name in enumerate(sheet_names):
        try:
            temp_df = pd.read_excel(stock_file, sheet_name=sheet_name, header=None)
            header_idx = 0
            found_header = False
            
            search_keywords = ['REFERENCE', 'DESIGNIATION', 'DATE', 'QUANTITE', 'REF']
            for row_idx, row in temp_df.head(50).iterrows():
                row_values = [str(val).upper().strip() for val in row.values if pd.notna(val)]
                matches = sum(1 for keyword in search_keywords if any(keyword in v for v in row_values))
                if matches >= 2:
                    header_idx = row_idx
                    found_header = True
                    break
            
            df = pd.read_excel(stock_file, sheet_name=sheet_name, header=header_idx)
            df = df.dropna(how="all")
            
            # Find reference and quantity columns
            ref_col = None
            qty_col = None
            
            custom_qty_col = None
            if quantity_cols is not None and idx < len(quantity_cols):
                custom_qty_col = quantity_cols[idx]
                if isinstance(custom_qty_col, list) and len(custom_qty_col) > 0:
                    custom_qty_col = custom_qty_col[0]
            
            for col in df.columns:
                col_upper = str(col).upper().strip()
                if ref_col is None and any(name.upper() in col_upper for name in stock_possible_col_names['ref']):
                    ref_col = col
                
                if qty_col is None:
                    if custom_qty_col is not None:
                        if custom_qty_col.upper() in col_upper or col_upper in custom_qty_col.upper():
                            qty_col = col
                    else:
                        if any(name.upper() in col_upper for name in stock_possible_col_names['quantity']):
                            qty_col = col
            
            if ref_col is None or qty_col is None:
                continue
            
            df[ref_col] = df[ref_col].astype('string')
            df[ref_col] = df[ref_col].str.strip()
            df[ref_col] = df[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
            
            df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0.0)
            df[qty_col] = df[qty_col].round(2)
            
            df_clean = df[[ref_col, qty_col]].copy()
            df_clean.columns = ['REFERENCE', 'QUANTITE']
            df_clean['SOURCE_SHEET'] = sheet_name
            
            all_stock_data.append(df_clean)
            
        except Exception as e:
            continue
    
    if all_stock_data:
        combined = pd.concat(all_stock_data, ignore_index=True)
        return combined
    else:
        return pd.DataFrame(columns=['REFERENCE', 'QUANTITE', 'SOURCE_SHEET'])


def process_atelier(atelier_key, stock_file_path, mov_file_path, month):
    """Process a single atelier and return matches/discrepancies"""
    
    if atelier_key not in sheet_args:
        return {'error': f'Unknown atelier: {atelier_key}', 'matches': [], 'discrepancies': []}
    
    args = sheet_args[atelier_key]
    
    try:
        possible_sheets = args['sheet_name']
        stock_sheet_config = args['stock_sheet']
        
        stock_sheets = stock_sheet_config['sheets']
        stock_quantity_cols = stock_sheet_config.get('quantity', None)
        
        if isinstance(possible_sheets, str):
            possible_sheets = [possible_sheets]
        
        # Check if paired mode
        is_paired_mode = (
            len(possible_sheets) > 0 and 
            isinstance(possible_sheets[0], list) and
            len(stock_sheets) > 0 and 
            isinstance(stock_sheets[0], list)
        )
        
        if is_paired_mode:
            # Paired mode - process each pair
            if len(possible_sheets) != len(stock_sheets):
                return {'error': f'Mismatched pairs: {len(possible_sheets)} mov vs {len(stock_sheets)} stock', 'matches': [], 'discrepancies': []}
            
            all_discrepancies = []
            all_matches = []
            
            for pair_idx, (mov_sheet_group, stock_sheet_group) in enumerate(zip(possible_sheets, stock_sheets)):
                pair_quantity_cols = None
                if stock_quantity_cols is not None and pair_idx < len(stock_quantity_cols):
                    pair_quantity_cols = stock_quantity_cols[pair_idx]
                
                # Read stock from this pair's sheets
                stock_df = read_stock_from_sheets(stock_file_path, stock_sheet_group, [pair_quantity_cols] if pair_quantity_cols else None)
                
                if stock_df.empty:
                    continue
                
                # Read and stack movement sheets
                mov_data_list = []
                for mov_sheet_name in mov_sheet_group:
                    try:
                        temp_df = pd.read_excel(mov_file_path, sheet_name=mov_sheet_name, header=None)
                        header_idx, found_header = find_header_row(temp_df, mov_possible_col_names)
                        
                        if found_header:
                            mov_single = pd.read_excel(mov_file_path, sheet_name=mov_sheet_name, header=header_idx)
                        else:
                            mov_single = pd.read_excel(mov_file_path, sheet_name=mov_sheet_name)
                        
                        # Find columns
                        found_cols = {}
                        for col_type, possible_names in mov_possible_col_names.items():
                            for name in possible_names:
                                if name in mov_single.columns:
                                    found_cols[col_type] = name
                                    break
                        
                        if 'date' not in found_cols:
                            for col in mov_single.columns:
                                if 'date' in str(col).lower():
                                    found_cols['date'] = col
                                    break
                        
                        if 'ref' not in found_cols or 'quantity' not in found_cols:
                            continue
                        
                        ref_col = found_cols['ref']
                        qty_col = found_cols['quantity']
                        
                        # Filter by date
                        if 'date' in found_cols:
                            date_col = found_cols['date']
                            parse_dates_normalized_eu(mov_single, date_col)
                            target_month = int(month)
                            target_year = 2025
                            
                            mask = (mov_single[date_col].dt.year < target_year) | \
                                   ((mov_single[date_col].dt.year == target_year) & (mov_single[date_col].dt.month <= target_month))
                            mov_single = mov_single[mask]
                        
                        # Clean
                        mov_single[ref_col] = mov_single[ref_col].astype('string').str.strip()
                        mov_single[ref_col] = mov_single[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
                        mov_single[qty_col] = pd.to_numeric(mov_single[qty_col], errors='coerce').fillna(0.0)
                        
                        mov_clean = mov_single[[ref_col, qty_col]].copy()
                        mov_clean.columns = ['Ref', 'Qty']
                        mov_data_list.append(mov_clean)
                        
                    except Exception:
                        continue
                
                if not mov_data_list:
                    continue
                
                mov_combined = pd.concat(mov_data_list, ignore_index=True)
                mov_agg = mov_combined.groupby('Ref')['Qty'].sum().reset_index()
                mov_agg.rename(columns={'Qty': 'Calc_Mov_Qty'}, inplace=True)
                mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)
                
                # Stock aggregation
                stock_agg = stock_df.groupby('REFERENCE')['QUANTITE'].sum().reset_index()
                stock_agg.rename(columns={'REFERENCE': 'Ref', 'QUANTITE': 'Stock_Qty'}, inplace=True)
                stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)
                
                # Compare
                comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
                comparison_df['Difference'] = comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']
                
                discrepancies = comparison_df[comparison_df['Difference'] != 0]
                matches = comparison_df[comparison_df['Difference'] == 0]
                
                all_discrepancies.extend(discrepancies.to_dict('records'))
                all_matches.extend(matches.to_dict('records'))
            
            return {
                'matches': all_matches,
                'discrepancies': sorted(all_discrepancies, key=lambda x: abs(x.get('Difference', 0)), reverse=True)
            }
        
        else:
            # Non-paired mode - simple processing
            # Read stock
            stock_df = read_stock_from_sheets(stock_file_path, stock_sheets, stock_quantity_cols)
            
            if stock_df.empty:
                return {'error': 'Could not read stock sheets', 'matches': [], 'discrepancies': []}
            
            # Read movement
            mov = None
            used_sheet = None
            
            for sheet in possible_sheets:
                try:
                    temp_df = pd.read_excel(mov_file_path, sheet_name=sheet, header=None)
                    header_idx, found_header = find_header_row(temp_df, mov_possible_col_names)
                    
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
                return {'error': f'Could not find sheets {possible_sheets}', 'matches': [], 'discrepancies': []}
            
            # Find columns
            found_cols = {}
            for col_type, possible_names in mov_possible_col_names.items():
                for name in possible_names:
                    if name in mov.columns:
                        found_cols[col_type] = name
                        break
            
            if 'date' not in found_cols:
                for col in mov.columns:
                    if 'date' in str(col).lower():
                        found_cols['date'] = col
                        break
            
            if 'ref' not in found_cols or 'quantity' not in found_cols:
                return {'error': f'Missing columns. Available: {mov.columns.tolist()}', 'matches': [], 'discrepancies': []}
            
            ref_col = found_cols['ref']
            qty_col = found_cols['quantity']
            
            # Filter by date
            if 'date' in found_cols:
                date_col = found_cols['date']
                parse_dates_normalized_eu(mov, date_col)
                target_month = int(month)
                target_year = 2025
                
                mask = (mov[date_col].dt.year < target_year) | \
                       ((mov[date_col].dt.year == target_year) & (mov[date_col].dt.month <= target_month))
                mov = mov[mask]
            
            # Clean
            object_columns = mov.select_dtypes(include=['object']).columns
            for col in object_columns:
                mov[col] = mov[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
            
            mov[ref_col] = mov[ref_col].astype('string')
            mov[ref_col] = mov[ref_col].str.replace(r'(?<=\d)\.(?=\d)', ',', regex=True)
            mov[qty_col] = pd.to_numeric(mov[qty_col], errors='coerce').fillna(0.0)
            
            # Aggregate
            mov_agg = mov.groupby(ref_col)[qty_col].sum().reset_index()
            mov_agg.rename(columns={ref_col: 'Ref', qty_col: 'Calc_Mov_Qty'}, inplace=True)
            mov_agg['Calc_Mov_Qty'] = mov_agg['Calc_Mov_Qty'].round(2)
            
            stock_agg = stock_df.groupby('REFERENCE')['QUANTITE'].sum().reset_index()
            stock_agg.rename(columns={'REFERENCE': 'Ref', 'QUANTITE': 'Stock_Qty'}, inplace=True)
            stock_agg['Stock_Qty'] = stock_agg['Stock_Qty'].round(2)
            
            # Compare
            comparison_df = pd.merge(stock_agg, mov_agg, on='Ref', how='outer').fillna(0)
            comparison_df['Difference'] = comparison_df['Stock_Qty'] - comparison_df['Calc_Mov_Qty']
            
            discrepancies = comparison_df[comparison_df['Difference'] != 0].sort_values(by='Difference', ascending=False)
            matches = comparison_df[comparison_df['Difference'] == 0]
            
            return {
                'matches': matches.to_dict('records'),
                'discrepancies': discrepancies.to_dict('records')
            }
        
    except Exception as e:
        return {'error': str(e), 'matches': [], 'discrepancies': []}


def process_all(stock_file, matched_files, month):
    """Process all matched files for Larbaa"""
    results = {}
    
    # Process each atelier
    for atelier_key, mov_file in matched_files.items():
        results[atelier_key] = process_atelier(atelier_key, stock_file['path'], mov_file['path'], month)
    
    return results
