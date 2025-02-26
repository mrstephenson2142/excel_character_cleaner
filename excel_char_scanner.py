import pandas as pd
import re
import sys
import os

def scan_excel_for_problematic_chars(file_path, problematic_chars=None):
    """
    Scan an Excel file for problematic characters and report their locations.
    
    Args:
        file_path: Path to the Excel file
        problematic_chars: List of problematic character patterns to search for (default: byte values 0x80-0xFF)
        
    Returns:
        List of tuples with (sheet_name, row, column, cell_value, matched_char)
    """
    if problematic_chars is None:
        # Default to searching for all extended ASCII characters (0x80-0xFF)
        problematic_chars = [chr(i) for i in range(0x80, 0x100)]
        # Add escape sequences that might appear in the string representation
        problematic_chars.extend([f"\\x{i:02x}" for i in range(0x80, 0x100)])
    
    results = []
    
    try:
        # Create Excel file reader
        excel_file = pd.ExcelFile(file_path)
        
        # Process each sheet
        for sheet_name in excel_file.sheet_names:
            print(f"Scanning sheet: {sheet_name}")
            
            # Read the sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Scan each cell
            for row_idx, row in enumerate(df.itertuples(), 1):
                for col_idx, cell_value in enumerate(row[1:], 1):
                    # Only process string values
                    if isinstance(cell_value, str):
                        # Check for problematic characters
                        for char in problematic_chars:
                            if char in cell_value or re.search(f"{re.escape(char)}", str(cell_value)):
                                # Convert to column name (A, B, C, etc.)
                                col_name = get_column_letter(col_idx)
                                col_header = df.columns[col_idx-1]
                                
                                results.append({
                                    'sheet': sheet_name,
                                    'row': row_idx,
                                    'column': col_name,
                                    'column_header': col_header,
                                    'cell_value': cell_value,
                                    'problematic_char': char,
                                    'hex_value': f"0x{ord(char):02x}" if len(char) == 1 else char
                                })
    
    except Exception as e:
        print(f"Error scanning Excel file: {e}")
        return []
    
    return results

def get_column_letter(col_idx):
    """
    Convert column index to Excel column letter (A, B, C, ..., Z, AA, AB, ...)
    """
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_char_scanner.py <excel_file_path> [specific_chars]")
        print("Example to scan for \\x81: python excel_char_scanner.py myfile.xlsx \\x81")
        return
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        return
    
    # If specific characters are provided, use them
    specific_chars = None
    if len(sys.argv) > 2:
        specific_chars = [sys.argv[2]]
        # Handle escape sequences properly
        specific_chars = [c.encode().decode('unicode_escape') if '\\' in c else c for c in specific_chars]
    
    results = scan_excel_for_problematic_chars(file_path, specific_chars)
    
    # Display results
    if not results:
        print(f"No problematic characters found in '{file_path}'.")
    else:
        print(f"\nFound {len(results)} instances of problematic characters:")
        print("-" * 80)
        for res in results:
            print(f"Sheet: {res['sheet']}")
            print(f"Location: Cell {res['column']}{res['row']} (Column Header: {res['column_header']})")
            print(f"Problematic Character: {res['hex_value']}")
            print(f"Cell Value: {res['cell_value']}")
            print("-" * 80)
            
    # Save results to CSV if there are any
    if results:
        output_file = f"{os.path.splitext(file_path)[0]}_char_scan_results.csv"
        pd.DataFrame(results).to_csv(output_file, index=False)
        print(f"Results saved to {output_file}")

if __name__ == "__main__":
    main()