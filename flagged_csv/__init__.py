"""
Flagged CSV - Convert XLSX files to CSV with visual formatting preserved as inline flags.

This library allows you to convert Excel files to CSV while preserving:
- Cell background colors as {{#RRGGBB}} flags
- Merged cell information as {{MG:XXXXXX}} flags
- Cell formatting (currency, dates, etc.)

Example:
    from flagged_csv import XlsxConverter
    
    converter = XlsxConverter()
    csv_content = converter.convert_to_csv(
        'data.xlsx',
        tab_name='Sheet1',
        include_colors=True,
        signal_merge=True
    )
"""

from .converter import XlsxConverter, XlsxConverterConfig
from .formatter import ExcelFormatter

__version__ = "0.1.0"
__all__ = ["XlsxConverter", "XlsxConverterConfig", "ExcelFormatter"]