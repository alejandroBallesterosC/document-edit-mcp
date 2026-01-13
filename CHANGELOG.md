# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.1] - 2026-01-13

### Added
- **Trash support** for file deletion: files can now be sent to the Windows Recycle Bin instead of permanent deletion
- `send2trash` dependency for cross-platform trash support

### Changed
- `delete_file` and `delete_directory` now require explicit confirmation parameter:
  - `"CORBEILLE"` - sends to Recycle Bin (recoverable)
  - `"SUPPRESSION DÉFINITIVE"` - permanent deletion (not recoverable)

## [0.3.0] - 2026-01-13

### Added
- **`delete_file`** - Delete a file with explicit user confirmation
- **`delete_directory`** - Delete an empty directory with explicit user confirmation
- File management operations section in capabilities

### Security
- All delete operations require explicit confirmation string to prevent accidental deletion

## [0.2.0] - 2025-12-01

### Added
- **`create_formatted_word_document`** - Create Word documents with rich formatting:
  - Headers and footers
  - Titles and subtitles with custom colors
  - Multiple heading levels
  - Paragraphs with alignment, bold, italic, custom colors
  - Bullet lists and numbered lists
  - Tables with header styling and alternating row colors
  - Key-value tables with first column highlighting
  - Page breaks and spacers
  - Support for `**bold**` markdown syntax in text
- **`read_word_document_structure`** - Analyze document structure:
  - Table count, column widths, row heights
  - Paragraph count
  - Header/footer detection
- **`compare_word_documents`** - Compare two documents structurally
- Comprehensive test suite for formatted documents

### Changed
- Updated documentation with JSON schema examples

### Important
- **Use `"sections"` NOT `"content"`** in JSON for `create_formatted_word_document`, otherwise the file will be empty!

## [0.1.0] - 2025-11-15

### Added
- Initial release
- Basic Word document operations (create, edit)
- Excel file operations (create, edit, convert from CSV)
- PDF operations (create, convert from Word)
- Text to Word conversion
