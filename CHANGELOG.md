# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.7] - 2026-08-31

### Fixed
- `max_rows`/`max_columns` now count **visible** rows/columns instead of
  physical spreadsheet ones. Previously the readers truncated at physical
  indices *before* dropping hidden rows/columns, so e.g. 50 hidden rows among
  the first 300 produced only 250 output rows. The physical read window is now
  widened to compensate (`_expand_limit`), including hidden runs that chain
  across the window boundary (e.g. rows 299-350 hidden with `max_rows=300`)
- Hidden-dimension detection is no longer capped at `max_rows`/`max_columns`
  (it is bounded by the sheet's used range instead), so hidden rows/columns
  just past the old cap are correctly accounted for

### Changed
- Sheets whose only hidden rows/columns lie beyond the visible window no longer
  take the slower formatting path; their output is unaffected by the hidden
  dimensions, so the fast engines are used

## [0.1.6] - 2026-07-27

### Fixed
- Hidden column ranges are now fully skipped. Excel stores consecutive hidden
  columns as a single `<col min="3" max="5" hidden="1"/>` element, which openpyxl
  exposes under one key (`C`); only the first column of the range used to be
  dropped, leaking the rest into the output
- Columns that survive hidden-column removal keep their **original** spreadsheet
  letters (e.g. `A,B,F,G`) instead of being renumbered from A, so the header now
  matches the `{l:XX}` coordinates produced by `--add-location`

### Changed
- Hidden row/column detection is shared by all reading paths (calamine, openpyxl,
  xlrd, direct openpyxl and the formatting path) via a single helper, so the
  behaviour no longer depends on which engine succeeds
- The `[0.1.1]` note below ("DataFrame columns are renamed to Excel-style
  letters") only holds when no column is hidden

## [0.1.3] - 2025-08-16

### Fixed
- Use workbook's actual indexed color palette instead of hardcoded values
- Indexed colors now correctly read from workbook._colors when available
- Falls back to openpyxl's COLOR_INDEX for standard colors
- Fixes issue where indexed color 9 was incorrectly showing as black (#000000) instead of white (#FFFFFF)

### Added
- Test case for indexed color handling

### Changed
- Removed hardcoded indexed color mapping that could be incorrect for custom palettes

## [0.1.2] - 2025-08-16

### Added
- Foreground color support with `{fc:#RRGGBB}` syntax for text/font colors
- Alternative background color syntax `{bc:#RRGGBB}` (backward-compatible with `{#RRGGBB}`)
- Intelligent black text contrast logic - only shows `{fc:#000000}` when background exists
- Separate ignore lists for foreground and background colors with sensible defaults
- New CLI options for granular color control:
  - `--include-colors`: Include both foreground and background colors
  - `--include-bg-colors`: Include background colors only
  - `--include-fg-colors`: Include foreground colors only
  - `--ignore-bg-colors`: Ignore specific background colors (default: #FFFFFF)
  - `--ignore-fg-colors`: Ignore specific foreground colors (default: #000000)
- Comprehensive tests for color handling and ignore defaults

### Changed
- `--ignore-colors` now applies default ignore lists for both foreground (#000000) and background (#FFFFFF)
- Color extraction logic optimized to only process foreground colors for non-empty cells
- Documentation updated with new color syntax examples

### Fixed
- Black text no longer appears unnecessarily on cells without backgrounds
- Color ignore functionality now properly handles separate foreground and background lists

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