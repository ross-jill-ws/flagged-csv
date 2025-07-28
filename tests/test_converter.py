"""
Tests for the flagged-csv converter.
"""

import pytest
import tempfile
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from flagged_csv import XlsxConverter, XlsxConverterConfig


def create_test_excel(file_path: Path, with_colors=False, with_merge=False):
    """Create a test Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    
    # Add some data
    ws['A1'] = 'Header1'
    ws['B1'] = 'Header2'
    ws['C1'] = 'Header3'
    
    ws['A2'] = 'Value1'
    ws['B2'] = 100
    ws['C2'] = 200.50
    
    # Add colors if requested
    if with_colors:
        ws['A1'].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        ws['B2'].fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    
    # Add merge if requested
    if with_merge:
        ws['D1'] = 'Merged Cell'
        ws.merge_cells('D1:F1')
    
    wb.save(file_path)
    return wb


class TestXlsxConverter:
    """Test the XLSX converter functionality."""
    
    def test_basic_conversion(self):
        """Test basic XLSX to CSV conversion."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test.xlsx"
            create_test_excel(xlsx_path)
            
            # Convert
            converter = XlsxConverter()
            result = converter.convert_to_csv(str(xlsx_path), 'TestSheet')
            
            # Check result
            assert 'Header1' in result
            assert 'Value1' in result
            assert '100' in result
    
    def test_color_extraction(self):
        """Test color flag extraction."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with colors
            xlsx_path = Path(temp_dir) / "test_colors.xlsx"
            create_test_excel(xlsx_path, with_colors=True)
            
            # Convert with colors
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path), 
                'TestSheet',
                include_colors=True
            )
            
            # Check for color flags
            assert '{{#FF0000}}' in result  # Red color
            assert '{{#00FF00}}' in result  # Green color
    
    def test_merge_detection(self):
        """Test merge cell detection."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with merged cells
            xlsx_path = Path(temp_dir) / "test_merge.xlsx"
            create_test_excel(xlsx_path, with_merge=True)
            
            # Convert with merge detection
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'TestSheet',
                signal_merge=True
            )
            
            # Check for merge flags
            assert '{{MG:' in result
            assert 'Merged Cell' in result
    
    def test_ignore_colors(self):
        """Test ignoring specific colors."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test_ignore.xlsx"
            wb = Workbook()
            ws = wb.active
            
            ws['A1'] = 'White BG'
            ws['A1'].fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            
            ws['B1'] = 'Red BG'
            ws['B1'].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            
            wb.save(xlsx_path)
            
            # Convert ignoring white
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                include_colors=True,
                ignore_colors='#FFFFFF'
            )
            
            # White should be ignored, red should be included
            assert '{{#FFFFFF}}' not in result
            assert '{{#FF0000}}' in result
    
    def test_invalid_sheet_name(self):
        """Test error handling for invalid sheet names."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test.xlsx"
            create_test_excel(xlsx_path)
            
            # Try to convert non-existent sheet
            converter = XlsxConverter()
            with pytest.raises(ValueError, match="Sheet 'InvalidSheet' not found"):
                converter.convert_to_csv(str(xlsx_path), 'InvalidSheet')
    
    def test_file_not_found(self):
        """Test error handling for missing files."""
        converter = XlsxConverter()
        with pytest.raises(FileNotFoundError):
            converter.convert_to_csv('/nonexistent/file.xlsx', 'Sheet1')
    
    def test_configuration(self):
        """Test converter configuration options."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test.xlsx"
            create_test_excel(xlsx_path)
            
            # Convert with custom config
            config = XlsxConverterConfig(
                header=False,
                keep_default_na=True
            )
            converter = XlsxConverter(config)
            result = converter.convert_to_csv(str(xlsx_path), 'TestSheet')
            
            # First line should be data, not headers
            lines = result.strip().split('\n')
            assert 'Header1' in lines[0]  # Headers are now data
    
    def test_output_formats(self):
        """Test different output formats."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test.xlsx"
            create_test_excel(xlsx_path)
            
            converter = XlsxConverter()
            
            # Test HTML output
            html_result = converter.convert_to_csv(
                str(xlsx_path),
                'TestSheet',
                output_format='html'
            )
            assert '<table' in html_result
            assert '<td>Header1</td>' in html_result
            
            # Test Markdown output
            md_result = converter.convert_to_csv(
                str(xlsx_path),
                'TestSheet',
                output_format='markdown'
            )
            assert '|' in md_result  # Markdown tables use pipes


if __name__ == '__main__':
    pytest.main([__file__, '-v'])