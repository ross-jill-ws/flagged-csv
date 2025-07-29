# Flagged CSV

Convert XLSX files to CSV while preserving visual formatting information as inline flags.
![ShowExcelConversionOriginalCsv_ManimCE_v0 19 0](https://github.com/user-attachments/assets/9241d467-ab47-4005-8438-3bad05864a6b)

Traditional XLSX → CSV conversion will cause the loss of format such as background cell colors or cell merge:
![ShowExcelConversionFlaggedCsv_ManimCE_v0 19 0](https://github.com/user-attachments/assets/2dee974a-d480-4f71-9404-10b121df388f)

The new Flagged CSV will attach important flags to the original cell value, keeping the colors/cell-merge with {} flags

## Overview

Flagged CSV is a Python library and command-line tool that converts Excel (XLSX) files to CSV format while preserving important visual information that would normally be lost in conversion:

- **Cell background colors** - Preserved as `{#RRGGBB}` flags
- **Merged cells** - Marked with `{MG:XXXXXX}` flags where XXXXXX is a unique identifier
- **Cell formatting** - Currency symbols, number formats, dates preserved as displayed in Excel

## Installation

### Via uv (Recommended)

```bash
# Clone the repository
git clone https://github.com/yourusername/flagged-csv.git
cd flagged-csv

# Sync dependencies with uv
uv sync

# Install in development mode
uv pip install -e .
```

### Via pip

```bash
pip install flagged-csv
```

## Quick Start

### Command Line Usage

```bash
# Basic conversion
flagged-csv input.xlsx -t Sheet1 > output.csv

# Include colors and merge information
flagged-csv input.xlsx -t Sheet1 --include-colors --signal-merge -o output.csv

# Preserve formatting and ignore white backgrounds
flagged-csv input.xlsx -t Sheet1 --preserve-formats --include-colors --ignore-colors "#FFFFFF"
```

### Python Library Usage

```bash
# Run Python scripts with uv
uv run python your_script.py
```

```python
# your_script.py
from flagged_csv import XlsxConverter

# Create converter instance
converter = XlsxConverter()

# Convert with all formatting options
csv_content = converter.convert_to_csv(
    'data.xlsx',
    tab_name='Sheet1',
    include_colors=True,
    signal_merge=True,
    preserve_formats=True,
    ignore_colors='#FFFFFF'
)

# Save to file
with open('output.csv', 'w') as f:
    f.write(csv_content)
```

## Flag Format Specification

### Color Flags
- Format: `{#RRGGBB}`
- Example: `Sales{#FF0000}` - "Sales" with red background
- Multiple flags can be combined: `100{#00FF00}{MG:123456}`

### Merge Flags  
- Format: `{MG:XXXXXX}` where XXXXXX is a 6-digit identifier
- All cells in a merged range share the same ID
- The first cell contains the actual value
- Subsequent cells contain only the merge flag

### Example Output

Given an Excel file with:
- Cell A1: "Total Sales" with blue background (#0000FF)
- Cells B1-D1: Merged cell containing "$1,000" with green background (#00FF00)

The CSV output would be:
```csv
Total Sales{#0000FF},$1000{#00FF00}{MG:384756},{MG:384756},{MG:384756}
```

## Configuration Options

### CLI Options

- `-t, --tab-name`: Sheet name to convert (required)
- `-o, --output`: Output file path (default: stdout)
- `--format`: Output format: csv, html, or markdown (default: csv)
- `--include-colors`: Include cell background colors
- `--signal-merge`: Include merged cell information
- `--preserve-formats`: Preserve number/date formatting
- `--ignore-colors`: Comma-separated hex colors to ignore
- `--no-header`: Exclude header row from output
- `--keep-na`: Keep NA values instead of converting to empty strings

### Python API Options

```python
from flagged_csv import XlsxConverter, XlsxConverterConfig

# Create converter with custom configuration
config = XlsxConverterConfig(
    keep_default_na=False,  # Convert NA to empty strings
    index=False,            # Don't include row index
    header=True             # Include column headers
)

converter = XlsxConverter(config)
```

## Advanced Usage

### Processing Multiple Sheets

Save this as `process_sheets.py`:

```python
from flagged_csv import XlsxConverter
import pandas as pd

converter = XlsxConverter()

# Process all sheets in a workbook
xl_file = pd.ExcelFile('multi_sheet.xlsx')
for sheet_name in xl_file.sheet_names:
    csv_content = converter.convert_to_csv(
        'multi_sheet.xlsx',
        tab_name=sheet_name,
        include_colors=True,
        signal_merge=True
    )
    
    with open(f'{sheet_name}.csv', 'w') as f:
        f.write(csv_content)
    print(f'Converted {sheet_name} -> {sheet_name}.csv')
```

Run with:
```bash
uv run python process_sheets.py
```

### Parsing Flagged CSV

Save this as `parse_flagged.py`:

```python
import re
import pandas as pd

def parse_flagged_csv(file_path):
    """Parse a flagged CSV file and extract values and formatting."""
    df = pd.read_csv(file_path, header=None)
    
    # Regular expressions for parsing flags
    color_pattern = r'{#([0-9A-Fa-f]{6})}'
    merge_pattern = r'{MG:(\d{6})}'
    
    # Extract clean values and formatting info
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell = str(df.iloc[row_idx, col_idx])
            
            # Extract color
            color_match = re.search(color_pattern, cell)
            if color_match:
                color = color_match.group(1)
                print(f"Cell ({row_idx},{col_idx}) has color #{color}")
            
            # Extract merge ID
            merge_match = re.search(merge_pattern, cell)
            if merge_match:
                merge_id = merge_match.group(1)
                print(f"Cell ({row_idx},{col_idx}) is part of merge group {merge_id}")
            
            # Get clean value (remove all flags)
            clean_value = re.sub(r'{[^}]+}', '', cell)
            df.iloc[row_idx, col_idx] = clean_value
    
    return df

# Example usage
if __name__ == "__main__":
    df = parse_flagged_csv('output.csv')
    print("\nCleaned data:")
    print(df)
```

Run with:
```bash
uv run python parse_flagged.py
```

### Working with Merged Cells

```python
def reconstruct_merged_cells(df):
    """Reconstruct merged cell ranges from flagged CSV."""
    merge_groups = {}
    
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell = str(df.iloc[row_idx, col_idx])
            
            # Find merge ID
            match = re.search(r'{MG:(\d{6})}', cell)
            if match:
                merge_id = match.group(1)
                if merge_id not in merge_groups:
                    merge_groups[merge_id] = []
                merge_groups[merge_id].append((row_idx, col_idx))
    
    # merge_groups now contains all cells belonging to each merge
    for merge_id, cells in merge_groups.items():
        print(f"Merge {merge_id}: {cells}")
```

## Output Formats

### CSV (Default)
Standard CSV format with flags appended to cell values.

### HTML
```python
html_output = converter.convert_to_csv(
    'data.xlsx',
    tab_name='Sheet1',
    output_format='html',
    include_colors=True
)
```

### Markdown
```python
markdown_output = converter.convert_to_csv(
    'data.xlsx', 
    tab_name='Sheet1',
    output_format='markdown',
    include_colors=True
)
```

## Error Handling

The library handles various error cases gracefully:

```python
try:
    csv_content = converter.convert_to_csv('data.xlsx', tab_name='InvalidSheet')
except ValueError as e:
    print(f"Sheet not found: {e}")
except FileNotFoundError as e:
    print(f"File not found: {e}")
```

## Performance Considerations

- The library uses multiple fallback engines (calamine, openpyxl, xlrd) for maximum compatibility
- Large files are processed efficiently with streaming where possible
- Color extraction uses caching to avoid repeated theme color lookups

## Testing

Run the test suite using uv:

```bash
# Run all tests
uv run pytest tests/test_converter.py

# Run tests with verbose output
uv run pytest tests/test_converter.py -v

# Run a specific test
uv run pytest tests/test_converter.py::TestXlsxConverter::test_color_extraction -v
```

## Development

```bash
# Set up development environment
uv sync

# Run the example script
uv run python example.py

# Run the CLI tool in development
uv run flagged-csv --help
```

## Requirements

- Python 3.11+
- pandas >= 2.0.0, < 2.2.0
- numpy < 2.0 (for compatibility)
- pydantic >= 2.0.0 (for configuration models)
- openpyxl >= 3.1.0
- python-calamine >= 0.2.0 (for robust Excel reading)
- xlrd >= 2.0.0 (for older Excel format support)
- click >= 8.0.0 (for CLI)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests to ensure everything works (`uv run pytest tests/`)
4. Commit your changes (`git commit -m 'Add amazing feature'`)
5. Push to the branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

This library is inspired by the need to preserve Excel's visual information during data processing pipelines, particularly for financial and business reporting applications where cell colors and merged cells convey important meaning.
