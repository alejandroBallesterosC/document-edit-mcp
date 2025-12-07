# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2024-12

### Added

- **`create_formatted_word_document`**: New function for creating professional Word documents with rich formatting support:
  - Title and subtitle with styling
  - Headers and footers
  - Multiple heading levels (1-4) with custom colors
  - Paragraphs with alignment, bold, italic, and color options
  - Rich text parsing (`**bold**` syntax within paragraphs)
  - Bullet lists and numbered lists
  - Data tables with header styling and alternating row colors
  - Key-value tables for forms and specifications
  - Page breaks and spacers
  - Custom column widths and row heights

- **`read_word_document_structure`**: New function to analyze document structure:
  - Table count and detailed structure (column widths via tblGrid, row heights)
  - Cell width analysis
  - Paragraph count
  - Header/footer detection

- **`compare_word_documents`**: New function to compare two documents structurally:
  - Table structure comparison
  - Difference reporting

- **Test suite**: Added comprehensive tests in `tests/test_formatted_documents.py`
  - 12 test cases covering all new functionality
  - Edge case handling

### Changed

- Updated README.md with full documentation for new functions
- Updated pyproject.toml with new version and metadata
- Bumped version to 0.2.0

### Internal

- Added helper functions for table formatting:
  - `_set_cell_shading`: Cell background colors
  - `_fix_cell_paragraph_spacing`: Proper paragraph spacing in cells
  - `_set_cell_borders`: Cell border styling
  - `_set_cell_width`: Cell width control
  - `_set_row_height`: Row height control (auto-fit or fixed)
  - `_set_table_grid`: Proper column widths via Word's tblGrid element
  - `_parse_rich_text`: Parse markdown-style bold markers

## [0.1.0] - 2024-11

### Added

- Initial release
- Basic Word document operations (create, edit, convert from txt)
- Basic Excel operations (create, edit, convert from csv)
- Basic PDF operations (create, convert from Word)
- MCP server integration with Claude Desktop
- Smithery installation support
