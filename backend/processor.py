"""
Stock Reconciliation Processor - Main Router
Routes processing requests to unit-specific processors
"""

import sys
import json
import os
import traceback


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


def match_files_to_ateliers(unit, files):
    """Match uploaded files to ateliers based on keywords"""
    ateliers = get_ateliers(unit)
    
    matched = {}
    unmatched_files = []
    
    for file_info in files:
        filename = file_info['name'].lower()
        found_match = False
        
        for atelier in ateliers:
            keyword = atelier.lower()
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


def process_files(unit, stock_file, matched_files, month):
    """Process files using the unit-specific processor"""
    log_debug(f"Processing unit: {unit}")
    log_debug(f"Stock file: {stock_file}")
    log_debug(f"Matched files: {json.dumps(list(matched_files.keys()))}")
    log_debug(f"Month: {month}")
    
    processor = get_unit_processor(unit)
    results = processor.process_all(stock_file, matched_files, month)
    
    return results


def export_results(results, output_dir):
    """Export results to CSV files"""
    import pandas as pd
    
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
        
        elif action == 'process':
            unit = request.get('unit')
            stock_file = request.get('stockFile')
            matched_files = request.get('matchedFiles', {})
            month = request.get('month')
            
            results = process_files(unit, stock_file, matched_files, month)
            response = {'success': True, 'results': results}
        
        elif action == 'export':
            results = request.get('results', {})
            output_dir = request.get('outputDir')
            
            os.makedirs(output_dir, exist_ok=True)
            exported = export_results(results, output_dir)
            response = {'success': True, 'exportedFiles': exported}
        
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
