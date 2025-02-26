## Cleaning Mode

The tool offers an interactive cleaning mode to fix problematic characters:

```bash
Would you like to clean the problematic characters?
This will create a new copy of the Excel file with the problematic characters handled.
Clean the file? (y/n): y
```

If you choose to clean the file, for each problematic character you'll be presented with options:

```
Cleaning cell Sheet1!C4
Current value: John Doe's Café
Problematic character: 0x81
Character: Non-printable - Unicode category: Cc (UNDEFINED)

Options:
1. Delete the character
2. Replace with custom text
3. Skip this cell
4. Skip all remaining cells
Choose an option (1-4): 
```

The tool will create:
1. A cleaned Excel file (`filename_cleaned.xlsx`)
2. A detailed log of all changes made (`filename_cleaning_log.txt`)# Excel Problematic Character Scanner

A Python utility for scanning Excel files to identify and locate problematic characters like `\x81` that can cause encoding issues.

## Overview

This tool helps identify problematic non-standard characters in Excel files that might cause issues when processing data. It reports the exact location (sheet, row, column) of each problematic character, as well as the specific position within the cell content, making it easy to find and fix these issues in your data. 

Additionally, it can identify whether characters are printable (human-readable) or not, and offers an interactive mode to clean problematic characters by either deleting them or replacing them with custom text.

## Installation

### Prerequisites

- Python 3.6 or higher

### Setup

1. Clone or download the repository containing the script
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Scan an Excel file for all problematic characters (extended ASCII range 0x80-0xFF):

```bash
python excel_char_scanner.py path/to/your/file.xlsx
```

### Scan for Specific Characters

To scan only for a specific character (e.g., `\x81`):

```bash
python excel_char_scanner.py path/to/your/file.xlsx \x81
```

To scan for multiple specific characters:

```bash
python excel_char_scanner.py path/to/your/file.xlsx "\x81\x82\x83"
```

## Output

The script produces two types of output:

1. **Console output** showing detailed information about each problematic character found:
   ```
   Sheet: Sheet1
   Location: Cell C4 (Column Header: Customer Name)
   Problematic Character: 0x81
   Character: Non-printable - Unicode category: Cc (UNDEFINED)
   Character Position(s) in Cell: 10
   Cell Value: John Doe's Café
   Context: ...ohn Doe's Café...
            ..........^
   ```

2. **CSV file** (saved in the same directory as the input file) containing all the results for further analysis.

## Example

```bash
# Scan a customer data file for the problematic character \x81
python excel_char_scanner.py customer_data.xlsx \x81

# Output will show all occurrences of \x81 in the file
# and save results to customer_data_char_scan_results.csv
```

## Troubleshooting

### Common Issues

1. **Missing dependencies**: Ensure you have installed all requirements with `pip install -r requirements.txt`

2. **File not found**: Verify the path to your Excel file is correct

3. **Permission denied**: Make sure you have read permissions for the Excel file

4. **Memory errors**: For very large Excel files, you might need to increase your system's available memory

## Advanced Usage

### Integration with Data Processing Pipelines

This tool can be imported as a module in your data processing scripts:

```python
from excel_char_scanner import scan_excel_for_problematic_chars

results = scan_excel_for_problematic_chars('data.xlsx', ['\x81', '\x82'])
```
