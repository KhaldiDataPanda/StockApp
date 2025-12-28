"""
Stock Reconciliation Processor - Main Router
Routes processing requests to unit-specific processors
"""

import sys
import json
import os
import traceback
import pandas as pd

# Fix Windows console encoding for Unicode
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def log_debug(msg):
    """Log debug message to stderr (captured by Electron for debugging)"""
    print(f"[DEBUG] {msg}", file=sys.stderr)


# Unit module mapping
UNIT_MODULES = {
    'Fath1': 'processor_fath1',
    'Fath2': 'processor_fath2',
    'Fath5': 'processor_fath5',
    'Larbaa': 'processor_larbaa',
    'Oran': 'processor_oran',
    'Fibre': 'processor_fibre',
    'Fath3': 'processor_fath3',
    'Mdoukal': 'processor_mdoukal',
    'Mags': 'processor_mags',
}


def get_unit_processor(unit):
    """Dynamically import and return the processor module for a unit"""
    if unit not in UNIT_MODULES:
        raise ValueError(f"Unknown unit: {unit}")
    
    module_name = UNIT_MODULES[unit]
    
    # Import the module
    try:
        module = __import__(module_name)
        return module
    except ImportError as e:
        log_debug(f"Failed to import {module_name}: {e}")
        raise ImportError(f"Processor module not found for unit: {unit}")


def get_ateliers(unit):
    """Get list of ateliers for a unit"""
    processor = get_unit_processor(unit)
    return processor.get_ateliers()


def get_sheet_args(unit):
    """Get sheet_args configuration for a unit"""
    processor = get_unit_processor(unit)
    return getattr(processor, 'sheet_args', {})


def get_mov_col_names(unit):
    """Get possible column names configuration for a unit"""
    processor = get_unit_processor(unit)
    return getattr(processor, 'mov_possible_col_names', {
        'date': ['Date', 'DATE', 'date'],
        'ref': ['Référence\nFournisseur', 'REFERENCE', 'reference', 'Référence', 'Référence\n'],
        'quantity': ['STOCK PV', 'STOCK', 'STOCKS', 'ST-P', 'ST-PV'],
    })


def match_files_to_ateliers(unit, files):
    """Match uploaded files to ateliers based on keywords"""
    ateliers = get_ateliers(unit)
    sheet_args = get_sheet_args(unit)
    
    matched = {}
    unmatched_files = []
    
    for file_info in files:
        filename = file_info['name'].lower()
        found_match = False
        
        for atelier in ateliers:
            # Allow units to specify a dedicated keyword for file matching
            keyword = str(sheet_args.get(atelier, {}).get('file_keyword', atelier)).lower()
            if keyword in filename:
                if atelier not in matched:
                    matched[atelier] = file_info
                    found_match = True
                    break
        
        if not found_match:
            unmatched_files.append(file_info)
    
    # Find unmatched ateliers
    unmatched_ateliers = [a for a in ateliers if a not in matched]
    
    return {
        'matched': matched,
        'unmatchedFiles': unmatched_files,
        'unmatchedAteliers': unmatched_ateliers
    }


def verify_files(unit, matched_files):
    """Verify sheet names and column names in matched files"""
    log_debug(f"Verifying files for unit: {unit}")
    
    sheet_args = get_sheet_args(unit)
    mov_col_names = get_mov_col_names(unit)
    
    verification_results = {}
    
    for atelier, file_info in matched_files.items():
        file_path = file_info.get('path') if isinstance(file_info, dict) else file_info
        filename = file_info.get('filename', os.path.basename(file_path)) if isinstance(file_info, dict) else os.path.basename(file_path)
        
        result = {
            'atelier': atelier,
            'filename': filename,
            'path': file_path,
            'valid': True,
            'errors': [],
            'availableSheets': [],
            'availableColumns': [],
            'expectedSheet': '',
            'detectedRefCol': None,
            'detectedQtyCol': None,
            'detectedDateCol': None
        }
        
        try:
            # Get expected sheet name from config
            if atelier in sheet_args:
                result['expectedSheet'] = sheet_args[atelier].get('sheet_name', '')
            
            # Read Excel file and get available sheets
            xl = pd.ExcelFile(file_path)
            result['availableSheets'] = xl.sheet_names
            
            # Check if expected sheet exists
            expected_sheet = result['expectedSheet']
            expected_candidates = []
            if isinstance(expected_sheet, (list, tuple)):
                expected_candidates = list(expected_sheet)
            elif expected_sheet not in (None, ''):
                expected_candidates = [expected_sheet]

            sheet_found = False
            first_found_sheet = None
            for candidate in expected_candidates:
                if candidate in xl.sheet_names:
                    sheet_found = True
                    first_found_sheet = candidate
                    break
            
            if not sheet_found and expected_candidates:
                result['valid'] = False
                result['errors'].append(f"Sheet '{expected_candidates[0]}' not found")
            
            # Try to read the sheet (use expected or first available)
            sheet_to_read = first_found_sheet if sheet_found else (xl.sheet_names[0] if xl.sheet_names else None)
            
            if sheet_to_read:
                # Read to find columns
                temp_df = pd.read_excel(file_path, sheet_name=sheet_to_read, header=None, nrows=20)
                
                # Try to find header row
                header_idx = 0
                for idx, row in temp_df.iterrows():
                    row_values = [str(val).strip() for val in row.values if pd.notna(val)]
                    if any(name in row_values for name in mov_col_names.get('date', [])):
                        header_idx = idx
                        break
                
                # Read with proper header
                df = pd.read_excel(file_path, sheet_name=sheet_to_read, header=header_idx)
                result['availableColumns'] = [str(col) for col in df.columns.tolist()]
                
                # Check for required columns
                found_ref = None
                for col_name in mov_col_names.get('ref', []):
                    if col_name in df.columns:
                        found_ref = col_name
                        break
                result['detectedRefCol'] = found_ref
                
                found_qty = None
                for col_name in mov_col_names.get('quantity', []):
                    if col_name in df.columns:
                        found_qty = col_name
                        break
                result['detectedQtyCol'] = found_qty
                
                found_date = None
                for col_name in mov_col_names.get('date', []):
                    if col_name in df.columns:
                        found_date = col_name
                        break
                result['detectedDateCol'] = found_date
                
                if not found_ref:
                    result['valid'] = False
                    result['errors'].append("Reference column not found")
                
                if not found_qty:
                    result['valid'] = False
                    result['errors'].append("Quantity column not found")
            
        except Exception as e:
            result['valid'] = False
            result['errors'].append(str(e))
            log_debug(f"Error verifying {filename}: {str(e)}")
        
        verification_results[atelier] = result
    
    return verification_results


def process_files(unit, stock_file, matched_files, month):
    """Process files using the unit-specific processor"""
    log_debug(f"Processing unit: {unit}")
    log_debug(f"Stock file: {stock_file}")
    log_debug(f"Matched files: {json.dumps(list(matched_files.keys()))}")
    log_debug(f"Month: {month}")
    
    processor = get_unit_processor(unit)
    results = processor.process_all(stock_file, matched_files, month)
    
    return results


def process_files_with_overrides(unit, stock_file, matched_files, month, overrides):
    """Process files with custom sheet/column overrides"""
    log_debug(f"Processing unit with overrides: {unit}")
    log_debug(f"Overrides: {json.dumps(overrides)}")
    
    processor = get_unit_processor(unit)
    
    # If the processor supports overrides, use them
    if hasattr(processor, 'process_all_with_overrides'):
        results = processor.process_all_with_overrides(stock_file, matched_files, month, overrides)
    else:
        # Fall back to regular processing but modify the sheet_args temporarily
        # This is a basic fallback - individual processors should implement their own
        log_debug("Processor doesn't support overrides, using regular processing")
        results = processor.process_all(stock_file, matched_files, month)
    
    return results


def export_results(results, output_dir):
    """Export results to CSV files"""
    
    exported = []
    
    for atelier, data in results.items():
        if atelier.startswith('_'):
            continue
            
        if 'error' in data:
            continue
        
        # Export matches
        if data.get('matches'):
            matches_df = pd.DataFrame(data['matches'])
            matches_path = os.path.join(output_dir, f'matches_{atelier}.csv')
            matches_df.to_csv(matches_path, index=False, encoding='utf-8-sig')
            exported.append(matches_path)
        
        # Export discrepancies
        if data.get('discrepancies'):
            disc_df = pd.DataFrame(data['discrepancies'])
            disc_path = os.path.join(output_dir, f'discrepancies_{atelier}.csv')
            disc_df.to_csv(disc_path, index=False, encoding='utf-8-sig')
            exported.append(disc_path)
    
    return exported


def export_to_excel(data, output_path):
    """Export data to Excel file"""
    try:
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False, engine='openpyxl')
        return True
    except Exception as e:
        log_debug(f"Error exporting to Excel: {str(e)}")
        raise e


def main():
    """Main entry point - receives JSON commands from Electron"""
    try:
        # Read JSON input from command line args (Electron uses this)
        if len(sys.argv) > 1:
            input_data = sys.argv[1]
        else:
            # Fallback to stdin
            input_data = sys.stdin.read()
        
        log_debug(f"Received input length: {len(input_data)}")
        
        request = json.loads(input_data)
        action = request.get('action')
        
        log_debug(f"Action: {action}")
        
        if action == 'get_ateliers':
            unit = request.get('unit')
            ateliers = get_ateliers(unit)
            response = {'success': True, 'ateliers': ateliers}
        
        elif action == 'match_files':
            unit = request.get('unit')
            files = request.get('files', [])
            result = match_files_to_ateliers(unit, files)
            response = {'success': True, **result}
        
        elif action == 'verify':
            unit = request.get('unit')
            matched_files = request.get('matchedFiles', {})
            result = verify_files(unit, matched_files)
            response = {'success': True, 'verification': result}
        
        elif action == 'process':
            unit = request.get('unit')
            stock_file = request.get('stockFile')
            matched_files = request.get('matchedFiles', {})
            month = request.get('month')
            overrides = request.get('overrides')
            
            if overrides:
                results = process_files_with_overrides(unit, stock_file, matched_files, month, overrides)
            else:
                results = process_files(unit, stock_file, matched_files, month)
            response = {'success': True, 'results': results}
        
        elif action == 'export':
            results = request.get('results', {})
            output_dir = request.get('outputDir')
            
            os.makedirs(output_dir, exist_ok=True)
            exported = export_results(results, output_dir)
            response = {'success': True, 'exportedFiles': exported}
        
        elif action == 'export_excel':
            data = request.get('data', [])
            output_path = request.get('outputPath')
            
            export_to_excel(data, output_path)
            response = {'success': True, 'exportedFile': output_path}
        
        else:
            response = {'success': False, 'error': f'Unknown action: {action}'}
        
        # Output JSON response
        print(json.dumps(response, ensure_ascii=False))
        
    except Exception as e:
        log_debug(f"Error: {traceback.format_exc()}")
        error_response = {
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }
        print(json.dumps(error_response, ensure_ascii=False))


if __name__ == '__main__':
    main()
