"""
Tests for the flagged-csv converter.
"""

import pytest
import tempfile
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

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
            assert '{#FF0000}' in result  # Red color
            assert '{#00FF00}' in result  # Green color
    
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
            assert '{MG:' in result
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
            assert '{#FFFFFF}' not in result
            assert '{#FF0000}' in result
    
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
            
            # Test with header=False (default)
            config = XlsxConverterConfig(
                header=False,
                keep_default_na=True
            )
            converter = XlsxConverter(config)
            result = converter.convert_to_csv(str(xlsx_path), 'TestSheet')
            
            # When header=False, we should not output DataFrame column headers (A, B, C)
            # but we should still include all rows of data
            lines = result.strip().split('\n')
            assert len(lines) == 2  # Both rows of data
            assert 'Header1' in lines[0]
            assert 'Value1' in lines[1]
            
            # Test with header=True
            config_with_header = XlsxConverterConfig(
                header=True,
                keep_default_na=True
            )
            converter_with_header = XlsxConverter(config_with_header)
            result_with_header = converter_with_header.convert_to_csv(str(xlsx_path), 'TestSheet')
            
            # When header=True, we should output DataFrame column headers (A, B, C)
            lines_with_header = result_with_header.strip().split('\n')
            assert len(lines_with_header) == 3  # Column headers + 2 data rows
            assert 'A,B,C' in lines_with_header[0]
            assert 'Header1' in lines_with_header[1]
    
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
            assert 'Header1' in html_result  # Header is in <th> not <td>
            assert '<td>Value1</td>' in html_result
            
            # Test Markdown output
            md_result = converter.convert_to_csv(
                str(xlsx_path),
                'TestSheet',
                output_format='markdown'
            )
            assert '|' in md_result  # Markdown tables use pipes
    
    def test_add_location(self):
        """Test adding location coordinates to cells."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file
            xlsx_path = Path(temp_dir) / "test_location.xlsx"
            wb = Workbook()
            ws = wb.active
            ws['A1'] = 'Header1'
            ws['B1'] = 'Header2'
            ws['A2'] = 'Value1'
            ws['B2'] = 100
            wb.save(xlsx_path)
            
            # Convert with location tags
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                add_location=True
            )
            
            # Check for location tags
            assert '{l:A1}' in result
            assert '{l:B1}' in result
            assert '{l:A2}' in result
            assert '{l:B2}' in result
    
    def test_keep_empty_lines(self):
        """Test keeping empty rows."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with empty rows
            xlsx_path = Path(temp_dir) / "test_empty.xlsx"
            wb = Workbook()
            ws = wb.active
            ws['A1'] = 'Row1'
            # Row 2 is empty
            ws['A3'] = 'Row3'
            # Row 4 is empty
            ws['A5'] = 'Row5'
            wb.save(xlsx_path)
            
            # Convert without keep_empty_lines (default)
            converter = XlsxConverter()
            result_no_empty = converter.convert_to_csv(str(xlsx_path), 'Sheet')
            lines_no_empty = result_no_empty.strip().split('\n')
            # Should have 3 data rows (no empty lines)
            assert len(lines_no_empty) == 3
            
            # Convert with keep_empty_lines
            result_with_empty = converter.convert_to_csv(
                str(xlsx_path), 'Sheet',
                keep_empty_lines=True
            )
            lines_with_empty = result_with_empty.strip().split('\n')
            # Should have 5 rows (including empty lines)
            assert len(lines_with_empty) == 5
    
    def test_black_text_contrast(self):
        """Test that black text is only included when there's a background color for contrast."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with various text/background combinations
            xlsx_path = Path(temp_dir) / "test_contrast.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # White text on dark background
            ws['A1'] = 'White on Dark'
            ws['A1'].font = Font(color='FFFFFF')
            ws['A1'].fill = PatternFill(start_color='8E1C02', end_color='8E1C02', fill_type='solid')
            
            # Black text on light background
            ws['B1'] = 'Black on Light'
            ws['B1'].font = Font(color='000000')
            ws['B1'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            # Black text with no background (should NOT include fc:#000000)
            ws['A2'] = 'Plain Black'
            ws['A2'].font = Font(color='000000')
            
            # Colored text with no background (should include the color)
            ws['B2'] = 'Red Text'
            ws['B2'].font = Font(color='FF0000')
            
            wb.save(xlsx_path)
            
            # Test 1: With default ignore lists (black fg ignored by default)
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                include_colors=True  # Include both fg and bg
            )
            
            lines = result.strip().split('\n')
            
            # Check white on dark background
            assert '{#8E1C02}' in lines[0]  # Dark background
            assert '{fc:#FFFFFF}' in lines[0]  # White text
            
            # Check black on light background (black text is ignored by default)
            assert '{#FFFF00}' in lines[0]  # Yellow background
            assert '{fc:#000000}' not in lines[0]  # Black text ignored by default
            
            # Check plain black text (no background)
            assert 'Plain Black' in lines[1]
            assert '{fc:#000000}' not in lines[1]  # Black text ignored by default
            
            # Check colored text (no background)
            assert 'Red Text' in lines[1]
            assert '{fc:#FF0000}' in lines[1]  # Red text should be included
            
            # Test 2: Without ignoring black text (to test contrast logic)
            result2 = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                include_colors=True,
                ignore_fg_colors=''  # Don't ignore any foreground colors
            )
            
            lines2 = result2.strip().split('\n')
            
            # Now black text on background should be included
            assert '{#FFFF00}' in lines2[0]  # Yellow background
            assert '{fc:#000000}' in lines2[0]  # Black text now included with background
            
            # Plain black text (no background) should still not be included due to contrast logic
            assert 'Plain Black' in lines2[1]
            assert '{fc:#000000}' not in lines2[1]  # Black text without background still not included
    
    def test_ignore_colors_defaults(self):
        """Test default ignore colors and custom ignore lists."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with various color combinations
            xlsx_path = Path(temp_dir) / "test_ignore.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # White background with black text (should be ignored by default)
            ws['A1'] = 'Normal'
            ws['A1'].fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            ws['A1'].font = Font(color='000000')
            
            # Colored background with colored text
            ws['B1'] = 'Colored'
            ws['B1'].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            ws['B1'].font = Font(color='00FF00')
            
            wb.save(xlsx_path)
            
            converter = XlsxConverter()
            
            # Test 1: With defaults (ignore white bg and black fg)
            result = converter.convert_to_csv(
                str(xlsx_path), 'Sheet',
                include_colors=True  # Both fg and bg
            )
            lines = result.strip().split('\n')
            # Normal cell should have no colors (white bg and black fg ignored)
            assert 'Normal,' in lines[0] or ',Normal' in lines[0]
            assert '{#FFFFFF}' not in lines[0]  # White bg ignored
            assert '{fc:#000000}' not in lines[0]  # Black fg ignored
            # Colored cell should have both colors
            assert '{#FF0000}' in lines[0]  # Red bg included
            assert '{fc:#00FF00}' in lines[0]  # Green fg included
            
            # Test 2: Override with empty ignore lists
            result = converter.convert_to_csv(
                str(xlsx_path), 'Sheet',
                include_colors=True,
                ignore_bg_colors='',  # Don't ignore any bg colors
                ignore_fg_colors=''   # Don't ignore any fg colors
            )
            lines = result.strip().split('\n')
            # Now white bg and black fg should be included
            assert '{#FFFFFF}' in lines[0]  # White bg now included
            assert '{fc:#000000}' in lines[0]  # Black fg now included
            
            # Test 3: Custom ignore lists
            result = converter.convert_to_csv(
                str(xlsx_path), 'Sheet',
                include_colors=True,
                ignore_bg_colors='#FF0000',  # Ignore red bg
                ignore_fg_colors='#00FF00'   # Ignore green fg
            )
            lines = result.strip().split('\n')
            # Red bg and green fg should be ignored
            assert '{#FF0000}' not in lines[0]  # Red bg ignored
            assert '{fc:#00FF00}' not in lines[0]  # Green fg ignored
            # White bg and black fg should be included (not in ignore list)
            assert '{#FFFFFF}' in lines[0]  # White bg included
            assert '{fc:#000000}' in lines[0]  # Black fg included (with background)
    
    def test_indexed_colors(self):
        """Test that indexed colors are correctly mapped."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with indexed colors
            xlsx_path = Path(temp_dir) / "test_indexed.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # Set cells with different indexed colors
            ws['A1'] = 'Red Background'
            ws['A1'].fill = PatternFill(patternType='solid', fgColor='00FF0000')  # Index 2 = Red
            
            ws['B1'] = 'Green Background'
            ws['B1'].fill = PatternFill(patternType='solid', fgColor='0000FF00')  # Index 3 = Green
            
            wb.save(xlsx_path)
            
            # Convert with colors
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                include_bg_colors=True,
                ignore_bg_colors=''  # Don't ignore any colors for this test
            )
            
            # Check that colors are extracted correctly
            assert '{#FF0000}' in result  # Red
            assert '{#00FF00}' in result  # Green
    
    def test_date_format_preservation(self):
        """Test that date formats are preserved correctly with --preserve-formats."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with various date formats
            xlsx_path = Path(temp_dir) / "test_dates.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # Set up dates with different formats
            from datetime import datetime
            test_date1 = datetime(2025, 10, 1)
            test_date2 = datetime(2025, 12, 31)
            
            ws['A1'] = 'Date Formats'
            ws['A2'] = test_date1
            ws['A3'] = test_date2
            
            # Set different date number formats
            ws['A2'].number_format = 'mm-dd-yy'  # Should become 10-01-25
            ws['A3'].number_format = 'yyyy-mm-dd'  # Should become 2025-12-31
            
            wb.save(xlsx_path)
            
            # Convert with preserve formats
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                preserve_formats=True
            )
            
            lines = result.strip().split('\n')
            
            # Check that dates are formatted according to their Excel format
            assert '10-01-25' in lines[1]  # mm-dd-yy format
            assert '2025-12-31' in lines[2]  # yyyy-mm-dd format
            
            # Ensure no platform-specific format issues (like %-d)
            assert '%-' not in result
            assert '%m' not in result
            assert '%d' not in result
    
    def test_hidden_rows_columns(self):
        """Test that hidden rows and columns are skipped by default."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with hidden rows and columns
            xlsx_path = Path(temp_dir) / "test_hidden.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # Set up test data
            ws['A1'] = 'Header1'
            ws['B1'] = 'Header2'
            ws['C1'] = 'Header3'
            
            ws['A2'] = 'Row2A'
            ws['B2'] = 'Row2B'
            ws['C2'] = 'Row2C'
            
            ws['A3'] = 'Row3A'  # This row will be hidden
            ws['B3'] = 'Row3B'
            ws['C3'] = 'Row3C'
            
            ws['A4'] = 'Row4A'
            ws['B4'] = 'Row4B'  # This column will be hidden
            ws['C4'] = 'Row4C'
            
            # Hide row 3 and column B
            ws.row_dimensions[3].hidden = True
            ws.column_dimensions['B'].hidden = True
            
            wb.save(xlsx_path)
            
            # Convert without any special flags
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet'
            )
            
            lines = result.strip().split('\n')
            
            # Should have 3 rows (header, row2, row4) but not row3
            assert len(lines) == 3
            
            # Row 3 should be skipped (hidden)
            assert 'Row3A' not in result
            assert 'Row3B' not in result
            assert 'Row3C' not in result
            
            # Column B should be skipped (hidden)
            assert 'Header2' not in result
            assert 'Row2B' not in result
            assert 'Row4B' not in result
            
            # But other cells should be present
            assert 'Header1' in result
            assert 'Header3' in result
            assert 'Row2A' in result
            assert 'Row2C' in result
            assert 'Row4A' in result
            assert 'Row4C' in result
            
            # Check structure: should be Header1,Header3 (not Header1,Header2,Header3 due to hidden column)
            expected_header = 'Header1,Header3'
            assert lines[0] == expected_header
    
    def test_hidden_column_range(self):
        """Test that a hidden column RANGE (C:E) is fully skipped.

        Excel stores consecutive hidden columns as a single
        <col min="3" max="5" hidden="1"/> element, which openpyxl exposes under the
        single key 'C'. Every column of the range must be dropped, and the surviving
        columns must keep their original spreadsheet letters (A,B,F,G).
        """
        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "test_hidden_range.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = 'Sheet1'

            for idx, value in enumerate(['AAA', 'BBB', 'CCC', 'DDD', 'EEE', 'FFF', 'GGG'], start=1):
                ws.cell(row=1, column=idx, value=value)

            # Hide columns C, D and E as a single range (as Excel itself writes it)
            hidden_range = ws.column_dimensions['C']
            hidden_range.min = 3
            hidden_range.max = 5
            hidden_range.hidden = True

            wb.save(xlsx_path)

            converter = XlsxConverter(XlsxConverterConfig(header=True))
            result = converter.convert_to_csv(str(xlsx_path), 'Sheet1')

            # Header keeps the ORIGINAL letters of the surviving columns
            assert result.strip().split('\n') == ['A,B,F,G', 'AAA,BBB,FFF,GGG']
            assert 'CCC' not in result
            assert 'DDD' not in result
            assert 'EEE' not in result

            # Location tags must agree with the header letters
            result_with_location = converter.convert_to_csv(
                str(xlsx_path), 'Sheet1', add_location=True
            )
            lines = result_with_location.strip().split('\n')
            assert lines[0] == 'A,B,F,G'
            assert lines[1] == 'AAA{l:A1},BBB{l:B1},FFF{l:F1},GGG{l:G1}'

    def test_hidden_single_column_preserves_letters(self):
        """Test that a single hidden column keeps the other columns' real letters."""
        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "test_hidden_single.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = 'Sheet1'

            for idx, value in enumerate(['AAA', 'BBB', 'CCC', 'DDD', 'EEE', 'FFF'], start=1):
                ws.cell(row=1, column=idx, value=value)

            # Hide only column C
            hidden_column = ws.column_dimensions['C']
            hidden_column.min = 3
            hidden_column.max = 3
            hidden_column.hidden = True

            wb.save(xlsx_path)

            converter = XlsxConverter(XlsxConverterConfig(header=True))
            result = converter.convert_to_csv(str(xlsx_path), 'Sheet1')

            # D, E and F keep their letters instead of being renumbered to C, D, E
            assert result.strip().split('\n') == ['A,B,D,E,F', 'AAA,BBB,DDD,EEE,FFF']
            assert 'CCC' not in result

    def test_max_rows_columns(self):
        """Test max rows and columns limits."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create test file with many rows and columns
            xlsx_path = Path(temp_dir) / "test_limits.xlsx"
            wb = Workbook()
            ws = wb.active
            
            # Create 10x10 grid
            for row in range(1, 11):
                for col in range(1, 11):
                    ws.cell(row=row, column=col, value=f"R{row}C{col}")
            
            wb.save(xlsx_path)
            
            # Convert with limits
            converter = XlsxConverter()
            result = converter.convert_to_csv(
                str(xlsx_path),
                'Sheet',
                max_rows=3,
                max_columns=4
            )
            
            # Parse result
            lines = result.strip().split('\n')
            # Should have only 3 rows
            assert len(lines) == 3
            
            # Each row should have only 4 columns
            for line in lines:
                cells = line.split(',')
                assert len(cells) == 4


if __name__ == '__main__':
    pytest.main([__file__, '-v'])