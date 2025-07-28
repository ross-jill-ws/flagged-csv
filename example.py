#!/usr/bin/env python3
"""
Example usage of the flagged-csv library.
"""

from flagged_csv import XlsxConverter, XlsxConverterConfig
import re
from pathlib import Path


def main():
    """Demonstrate various features of the flagged-csv library."""
    
    # Example 1: Basic conversion
    print("=== Example 1: Basic Conversion ===")
    converter = XlsxConverter()
    
    # Note: You'll need to provide your own Excel file
    excel_file = "example_data.xlsx"
    
    if Path(excel_file).exists():
        try:
            csv_content = converter.convert_to_csv(
                excel_file,
                tab_name="Sheet1"
            )
            print("Basic CSV output:")
            print(csv_content[:200] + "..." if len(csv_content) > 200 else csv_content)
            print()
        except Exception as e:
            print(f"Error: {e}")
    else:
        print(f"Please provide {excel_file} to run this example")
        print()
    
    # Example 2: Full formatting
    print("=== Example 2: Full Formatting Options ===")
    
    # Configure the converter
    config = XlsxConverterConfig(
        keep_default_na=False,  # Convert NA to empty strings
        header=True,            # Include headers
        index=False            # Don't include row index
    )
    
    converter_with_config = XlsxConverter(config)
    
    # Simulated conversion (would need actual file)
    print("""
    # This would convert with all formatting options:
    csv_with_formatting = converter.convert_to_csv(
        'financial_report.xlsx',
        tab_name='Q4 Results',
        include_colors=True,      # Include {{#RRGGBB}} color flags
        signal_merge=True,        # Include {{MG:XXXXXX}} merge flags
        preserve_formats=True,    # Keep $, %, dates as formatted
        ignore_colors='#FFFFFF'   # Ignore white backgrounds
    )
    """)
    
    # Example 3: Parsing flagged CSV
    print("=== Example 3: Parsing Flagged CSV ===")
    
    # Sample flagged CSV content
    sample_flagged_csv = """Revenue{{#00FF00}},$1000{{#00FF00}}{{MG:123456}},{{MG:123456}},Q4 Total
Expenses{{#FF0000}},$800{{#FF0000}}{{MG:789012}},{{MG:789012}},Q4 Total
Profit{{#0000FF}},$200{{#0000FF}},,Q4 Total"""
    
    print("Sample flagged CSV:")
    print(sample_flagged_csv)
    print()
    
    # Parse the flags
    lines = sample_flagged_csv.strip().split('\n')
    for i, line in enumerate(lines):
        cells = line.split(',')
        print(f"\nRow {i + 1}:")
        
        for j, cell in enumerate(cells):
            # Extract color
            color_match = re.search(r'{{#([0-9A-Fa-f]{6})}}', cell)
            color = f"#{color_match.group(1)}" if color_match else "No color"
            
            # Extract merge ID
            merge_match = re.search(r'{{MG:(\d{6})}}', cell)
            merge_id = merge_match.group(1) if merge_match else "Not merged"
            
            # Get clean value
            clean_value = re.sub(r'{{[^}]+}}', '', cell)
            
            print(f"  Cell {chr(65+j)}{i+1}: '{clean_value}' | Color: {color} | Merge: {merge_id}")
    
    # Example 4: Working with multiple sheets
    print("\n=== Example 4: Multiple Sheets ===")
    print("""
    # Process all sheets in a workbook:
    
    import pandas as pd
    
    xl_file = pd.ExcelFile('multi_sheet_workbook.xlsx')
    for sheet_name in xl_file.sheet_names:
        csv_content = converter.convert_to_csv(
            'multi_sheet_workbook.xlsx',
            tab_name=sheet_name,
            include_colors=True,
            signal_merge=True
        )
        
        # Save each sheet to a separate CSV
        output_file = f'{sheet_name.replace(" ", "_")}.csv'
        with open(output_file, 'w') as f:
            f.write(csv_content)
        print(f'Converted {sheet_name} -> {output_file}')
    """)
    
    # Example 5: Different output formats
    print("\n=== Example 5: Output Formats ===")
    print("""
    # HTML output
    html_output = converter.convert_to_csv(
        'report.xlsx',
        tab_name='Summary',
        output_format='html',
        include_colors=True
    )
    
    # Markdown output
    markdown_output = converter.convert_to_csv(
        'report.xlsx',
        tab_name='Summary', 
        output_format='markdown',
        include_colors=True
    )
    """)


if __name__ == "__main__":
    main()