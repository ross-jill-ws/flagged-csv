# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2024-01-XX

### Added
- Initial release of flagged-csv
- XLSX to CSV conversion with visual formatting preservation
- Cell background color flags `{#RRGGBB}`
- Merged cell flags `{MG:XXXXXX}`
- Cell formatting preservation (currency, dates, etc.)
- CLI tool for command-line usage
- Support for multiple output formats (CSV, HTML, Markdown)
- Robust file reading with multiple engine fallbacks
- Comprehensive test suite
- Documentation and examples

### Features
- Process Excel files while preserving visual information for AI processing
- Color pattern detection and categorization
- Time range detection from merged cells
- Compatible with Python 3.11+