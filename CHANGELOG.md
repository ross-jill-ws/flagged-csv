# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.1] - 2025-08-15

### Added
- Cell location flags `{l:CellRef}` to preserve original Excel cell coordinates
- `--add-location` CLI option to include cell coordinates in output
- `--keep-empty-lines` CLI option to preserve empty rows for maintaining structure
- `--max-rows` and `--max-columns` CLI options for processing limits (defaults: 300/100)
- `keep_empty_lines` configuration option in `XlsxConverterConfig`
- `add_location` configuration option in `XlsxConverterConfig`
- `max_rows` and `max_columns` parameters to `convert_to_csv()` method
- `_remove_empty_rows()` method to filter out empty rows when needed
- Enhanced `_trim_trailing_empty_rows()` method to remove trailing empty content

### Changed
- Default value for `header` configuration changed from `True` to `False` to match vendor-python implementation
- All Excel reading methods now treat all rows as data (header=None) by default
- DataFrame columns are renamed to Excel-style letters (A, B, C...) for consistency
- Empty rows are now removed by default (can be preserved with `keep_empty_lines=True`)

### Fixed
- NumPy 2.x compatibility issues by pinning to numpy<2.0
- Test suite now correctly handles header behavior
- Empty row handling in various Excel reading engines

## [0.1.0] - 2024-12-XX

### Added
- Initial release of flagged-csv
- XLSX to CSV conversion with visual formatting preservation
- Cell background color flags `{#RRGGBB}`
- Merged cell flags `{MG:XXXXXX}`
- Cell formatting preservation (currency, dates, etc.)
- CLI tool for command-line usage
- Support for multiple output formats (CSV, HTML, Markdown)
- Robust file reading with multiple engine fallbacks (calamine, openpyxl, xlrd)
- Theme color extraction and tint calculations
- Ignore specific colors functionality
- Comprehensive test suite
- Documentation and examples
- AI integration guide with flagged-csv.prompt.md

### Features
- Process Excel files while preserving visual information for AI processing
- Color pattern detection and categorization
- Time range detection from merged cells
- Compatible with Python 3.11+
- UV package manager support