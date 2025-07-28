"""
Main XLSX to Flagged CSV converter implementation.
"""

from typing import Optional, Literal, Dict, Any
from pathlib import Path
import random
import warnings
import colorsys
import zipfile
import xml.etree.ElementTree as etree

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

from .formatter import ExcelFormatter


class XlsxConverterConfig(BaseModel):
    """Configuration for XLSX conversion."""
    
    keep_default_na: bool = Field(
        default=False,
        description="Whether to keep default NA values in pandas"
    )
    index: bool = Field(
        default=False,
        description="Whether to include index in CSV output"
    )
    header: bool = Field(
        default=True,
        description="Whether to include header in CSV output"
    )


class XlsxConverter:
    """Convert XLSX files to CSV format with optional formatting flags."""
    
    def __init__(self, config: Optional[XlsxConverterConfig] = None):
        """
        Initialize the converter.
        
        Args:
            config: Optional configuration for the converter
        """
        self.config = config or XlsxConverterConfig()
        self._patch_openpyxl_colors()
        self._cached_theme_colors = None
    
    def _patch_openpyxl_colors(self):
        """Patch openpyxl to handle color validation issues."""
        import openpyxl.styles.colors
        original_rgb_set = openpyxl.styles.colors.RGB.__set__
        
        def patched_rgb_set(self, instance, value):
            if value is not None:
                if isinstance(value, str):
                    value = ''.join(c for c in value if c in '0123456789ABCDEFabcdef')
                    if len(value) == 6:
                        value = 'FF' + value
                    elif len(value) < 6:
                        value = value.ljust(8, '0')
                    elif len(value) > 8:
                        value = value[:8]
            try:
                original_rgb_set(self, instance, value)
            except ValueError:
                instance._rgb = None
        
        openpyxl.styles.colors.RGB.__set__ = patched_rgb_set
    
    def convert_to_csv(
        self,
        input_file_path: str,
        tab_name: str,
        output_format: Literal["csv", "html", "markdown"] = "csv",
        include_colors: bool = False,
        signal_merge: bool = False,
        preserve_formats: bool = False,
        ignore_colors: Optional[str] = None
    ) -> str:
        """
        Convert an XLSX file to CSV with optional formatting flags.
        
        Args:
            input_file_path: Path to the XLSX file
            tab_name: Name of the sheet to convert
            output_format: Format to convert to (csv, html, or markdown)
            include_colors: Whether to include cell background colors as {{#RRGGBB}} flags
            signal_merge: Whether to include merge information as {{MG:XXXXXX}} flags
            preserve_formats: Whether to preserve cell formatting (e.g., $500 instead of 500)
            ignore_colors: Comma-separated string of hex colors to ignore (e.g., "#FFFFFF,#000000")
            
        Returns:
            Formatted string in the requested format
            
        Raises:
            FileNotFoundError: If the input file doesn't exist
            ValueError: If the tab name doesn't exist in the XLSX file
        """
        # Verify file exists
        path = Path(input_file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {input_file_path}")
        if path.suffix.lower() not in ['.xlsx', '.xls']:
            raise ValueError(f"File must be an Excel file (xlsx/xls), got: {path.suffix}")
        
        try:
            # Determine whether to use special formatting
            if include_colors or signal_merge or preserve_formats:
                df = self._read_excel_with_formatting(
                    input_file_path, tab_name, include_colors, 
                    signal_merge, preserve_formats, ignore_colors
                )
            else:
                df = self._read_excel_with_fallback(
                    input_file_path, tab_name, self.config.keep_default_na
                )
            
            # Convert to requested format
            if output_format == "csv":
                return df.to_csv(index=self.config.index, header=self.config.header)
            elif output_format == "html":
                return df.to_html(index=self.config.index, header=self.config.header)
            elif output_format == "markdown":
                return df.to_markdown(index=self.config.index)
            
        except ValueError as e:
            if "No sheet named" in str(e) or "not found" in str(e):
                raise ValueError(f"Sheet '{tab_name}' not found in {path.name}")
            raise
    
    def _read_excel_with_fallback(self, file_path: str, sheet_name: str, keep_default_na: bool) -> pd.DataFrame:
        """Read Excel file with multiple engine fallbacks."""
        exceptions = []
        
        # Try calamine engine first
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name, 
                               keep_default_na=keep_default_na, engine='calamine')
        except Exception as e:
            exceptions.append(f"Calamine: {e}")
        
        # Try openpyxl
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name, 
                               keep_default_na=keep_default_na, engine='openpyxl')
        except Exception as e:
            exceptions.append(f"Openpyxl: {e}")
        
        # Try xlrd
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name, 
                               keep_default_na=keep_default_na, engine='xlrd')
        except Exception as e:
            exceptions.append(f"Xlrd: {e}")
        
        # Direct openpyxl reading
        try:
            warnings.filterwarnings('ignore')
            wb = load_workbook(file_path, data_only=True, keep_links=False)
            ws = wb[sheet_name]
            
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)
            
            if data:
                df = pd.DataFrame(data[1:], columns=data[0])
                if not keep_default_na:
                    df = df.fillna('')
                return df
            return pd.DataFrame()
            
        except Exception as e:
            exceptions.append(f"Direct openpyxl: {e}")
        
        # Check for sheet not found errors
        for exc in exceptions:
            if "not found" in exc.lower() or "does not exist" in exc.lower():
                raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")
        
        raise Exception(f"Unable to read Excel file: {'; '.join(exceptions)}")
    
    def _extract_theme_colors(self, file_path: str) -> Dict[int, str]:
        """Extract theme colors from XLSX file."""
        theme_colors = {}
        
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                if 'xl/theme/theme1.xml' in zip_file.namelist():
                    theme_xml = zip_file.read('xl/theme/theme1.xml')
                    root = etree.fromstring(theme_xml)
                    
                    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                    color_scheme = root.find('.//a:clrScheme', ns)
                    
                    if color_scheme is not None:
                        color_map = [
                            ('lt1', 0), ('dk1', 1), ('lt2', 2), ('dk2', 3),
                            ('accent1', 4), ('accent2', 5), ('accent3', 6),
                            ('accent4', 7), ('accent5', 8), ('accent6', 9),
                            ('hlink', 10), ('folHlink', 11)
                        ]
                        
                        for color_name, idx in color_map:
                            elem = color_scheme.find(f'.//a:{color_name}', ns)
                            if elem is not None:
                                srgb = elem.find('.//a:srgbClr', ns)
                                sys_color = elem.find('.//a:sysClr', ns)
                                
                                if srgb is not None:
                                    theme_colors[idx] = srgb.get('val')
                                elif sys_color is not None:
                                    theme_colors[idx] = sys_color.get('lastClr', '000000')
        except Exception:
            pass
        
        return theme_colors
    
    def _apply_tint(self, rgb_hex: str, tint: float) -> str:
        """Apply tint to a color."""
        if not tint or tint == 0:
            return rgb_hex
        
        # Convert hex to RGB
        r = int(rgb_hex[0:2], 16) / 255.0
        g = int(rgb_hex[2:4], 16) / 255.0
        b = int(rgb_hex[4:6], 16) / 255.0
        
        # Convert to HSL
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        
        # Apply tint
        if tint < 0:
            l = l * (1 + tint)  # Make darker
        else:
            l = l + (1 - l) * tint  # Make lighter
        
        # Convert back to RGB
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        
        return f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
    
    def _read_excel_with_formatting(
        self,
        file_path: str,
        sheet_name: str,
        include_colors: bool = True,
        signal_merge: bool = False,
        preserve_formats: bool = False,
        ignore_colors: Optional[str] = None
    ) -> pd.DataFrame:
        """Read Excel file and extract formatting information."""
        # Parse colors to ignore
        colors_to_ignore = set()
        if ignore_colors:
            for color in ignore_colors.split(','):
                color = color.strip().upper()
                if color.startswith('#'):
                    color = color[1:]
                colors_to_ignore.add(color)
        
        wb = load_workbook(file_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")
        
        ws = wb[sheet_name]
        
        # Clear cached theme colors
        self._cached_theme_colors = None
        
        # Build merge map
        merge_map = {}
        if signal_merge and ws.merged_cells:
            for merged_range in ws.merged_cells.ranges:
                merge_id = str(random.randint(100000, 999999))
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                        merge_map[(row, col)] = merge_id
        
        # Process all cells
        processed_data = []
        
        for row_idx, row in enumerate(ws.iter_rows(), 1):
            row_data = []
            has_content = False
            
            for col_idx, cell in enumerate(row, 1):
                value = cell.value
                if value is not None:
                    has_content = True
                
                # Apply formatting if requested
                if preserve_formats and value is not None:
                    try:
                        if cell.number_format and cell.number_format != 'General':
                            value = ExcelFormatter.format_value(value, cell.number_format)
                    except:
                        pass
                
                formatting_parts = []
                
                # Extract color if requested
                if include_colors:
                    color_hex = self._extract_cell_color(cell, file_path)
                    if color_hex:
                        color_to_check = color_hex[1:] if color_hex.startswith('#') else color_hex
                        if color_to_check.upper() not in colors_to_ignore:
                            formatting_parts.append(f"{{{color_hex}}}")
                
                # Add merge info if requested
                if signal_merge and (cell.row, cell.column) in merge_map:
                    merge_id = merge_map[(cell.row, cell.column)]
                    formatting_parts.append(f"{{MG:{merge_id}}}")
                
                # Combine value and formatting
                if value is not None:
                    cell_value = str(value) + ''.join(formatting_parts)
                elif formatting_parts:
                    cell_value = ''.join(formatting_parts)
                else:
                    cell_value = None
                
                row_data.append(cell_value)
            
            if has_content:
                processed_data.append(row_data)
        
        if not processed_data:
            return pd.DataFrame()
        
        # Create DataFrame
        num_cols = ws.max_column
        headers = [get_column_letter(i) for i in range(1, num_cols + 1)]
        
        # Ensure all rows have same number of columns
        max_cols = max(len(row) for row in processed_data)
        for row in processed_data:
            while len(row) < max_cols:
                row.append(None)
        
        df = pd.DataFrame(processed_data, columns=headers[:max_cols])
        
        if not self.config.keep_default_na:
            df = df.fillna('')
        
        return df
    
    def _extract_cell_color(self, cell, file_path: str) -> Optional[str]:
        """Extract color from a cell."""
        if not cell.fill or cell.fill.patternType != 'solid':
            return None
        
        color = cell.fill.fgColor or cell.fill.start_color
        if not color:
            return None
        
        try:
            if hasattr(color, 'type'):
                if color.type == 'theme' and hasattr(color, 'theme') and color.theme is not None:
                    # Load theme colors
                    if self._cached_theme_colors is None:
                        self._cached_theme_colors = self._extract_theme_colors(file_path)
                    
                    # Default theme colors if extraction fails
                    if not self._cached_theme_colors:
                        self._cached_theme_colors = {
                            0: "FFFFFF", 1: "000000", 2: "E7E6E6", 3: "44546A",
                            4: "5B9BD5", 5: "ED7D31", 6: "A5A5A5", 7: "FFC000",
                            8: "4472C4", 9: "70AD47"
                        }
                    
                    base_color = self._cached_theme_colors.get(color.theme, "000000")
                    
                    # Apply tint if present
                    tint = getattr(color, 'tint', 0)
                    if tint and tint != 0:
                        final_color = self._apply_tint(base_color, tint)
                        return f"#{final_color}"
                    return f"#{base_color}"
                
                elif color.type == 'indexed' and hasattr(color, 'indexed'):
                    # Standard indexed colors
                    indexed_colors = {
                        0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00",
                        4: "#0000FF", 5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF",
                        # ... (abbreviated for brevity)
                    }
                    return indexed_colors.get(color.indexed, "#000000")
                
                elif color.type == 'rgb' or not hasattr(color, 'type'):
                    if hasattr(color, 'rgb'):
                        rgb = color.rgb
                        if isinstance(rgb, str) and 'Values must be' not in rgb:
                            if len(rgb) == 8:
                                return f"#{rgb[2:]}"
                            elif len(rgb) == 6:
                                return f"#{rgb}"
        except:
            pass
        
        return None