[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/alejandroballesterosc-document-edit-mcp-badge.png)](https://mseep.ai/app/alejandroballesterosc-document-edit-mcp)

# Claude Document MCP Server
[![smithery badge](https://smithery.ai/badge/@alejandroBallesterosC/document-edit-mcp)](https://smithery.ai/server/@alejandroBallesterosC/document-edit-mcp)

A Model Context Protocol (MCP) server that allows Claude Desktop to perform document operations on Microsoft Word, Excel, and PDF files.

## Features

### Microsoft Word Operations
- **Create formatted Word documents** with rich styling (tables, colors, headers/footers, lists)
- Create simple Word documents from text
- Edit existing Word documents (add/edit/delete paragraphs and headings)
- Convert text files (.txt) to Word documents
- **Analyze document structure** (tables, columns, rows)
- **Compare two documents** structurally

### Excel Operations
- Create new Excel spreadsheets from JSON or CSV-like text
- Edit existing Excel files (update cells, ranges, add/delete rows, columns, sheets)
- Convert CSV files to Excel

### PDF Operations
- Create new PDF files from text
- Convert Word documents to PDF files

## Setup

This MCP server requires Python 3.10 or higher.

### Installing via Smithery

To install Claude Document MCP Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@alejandroBallesterosC/document-edit-mcp):

```bash
npx -y @smithery/cli install @alejandroBallesterosC/document-edit-mcp --client claude
```

### Automatic Setup (Recommended)

Run the setup script to automatically install dependencies and configure for Claude Desktop:

```bash
git clone https://github.com/alejandroBallesterosC/document-edit-mcp
cd document-edit-mcp
./setup.sh
```

This will:
1. Create a virtual environment
2. Install required dependencies
3. Configure the server for Claude Desktop
4. Create necessary directories

### Manual Setup

If you prefer to set up manually:

1. Install dependencies:

```bash
cd claude-document-mcp
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -e .
```

2. Configure Claude Desktop:

Copy the `claude_desktop_config.json` file to:
- **Mac**: `~/Library/Application Support/Claude/`
- **Windows**: `%APPDATA%\Claude\`

3. Restart Claude Desktop

## Model Context Protocol Integration

This server follows the Model Context Protocol specification to provide document manipulation capabilities for Claude Desktop:

- **Tools**: Provides manipulations functions for Word, Excel, and PDF operations
- **Resources**: Provides information about capabilities
- **Prompts**: (none currently implemented)

## API Reference

### Microsoft Word

#### Create a Formatted Word Document (NEW)

Creates professional Word documents with rich formatting support.

```python
create_formatted_word_document(filepath: str, document_data: str) -> Dict
```

**Parameters:**
- `filepath`: Path where to save the document
- `document_data`: JSON string containing document structure

**Supported section types:**
- `heading`: Headings level 1-4 with custom colors
- `paragraph`: Text with bold, italic, color, alignment options
- `bullet_list`: Bulleted lists
- `numbered_list`: Numbered lists
- `table`: Data tables with header styling and alternating row colors
- `key_value_table`: Two-column tables ideal for forms/specs
- `page_break`: Insert page breaks
- `spacer`: Add vertical spacing

**Example:**
```json
{
  "title": "Project Report",
  "subtitle": "Q4 2024",
  "header": "Confidential",
  "footer": "Page 1",
  "sections": [
    {
      "type": "heading",
      "level": 1,
      "text": "Executive Summary",
      "color": "1F4E79"
    },
    {
      "type": "paragraph",
      "text": "This report covers **key findings** from Q4.",
      "alignment": "justify"
    },
    {
      "type": "table",
      "headers": ["Metric", "Value", "Change"],
      "rows": [
        ["Revenue", "$1.2M", "+15%"],
        ["Users", "50,000", "+22%"]
      ],
      "header_bg_color": "1F4E79",
      "header_text_color": "FFFFFF",
      "alt_row_color": "F2F2F2"
    },
    {
      "type": "key_value_table",
      "rows": [
        ["Project", "Alpha"],
        ["Status", "On Track"],
        ["Owner", "John Doe"]
      ],
      "first_col_bg_color": "D6E3F0",
      "first_cell_bg_color": "1F4E79",
      "first_cell_text_color": "FFFFFF"
    },
    {
      "type": "bullet_list",
      "items": ["Increased market share", "Launched 3 new products", "Expanded to 2 new regions"]
    }
  ]
}
```

**Rich text support:** Use `**text**` for bold within paragraphs.

#### Read Document Structure (NEW)

Analyzes the internal structure of a Word document.

```python
read_word_document_structure(filepath: str) -> Dict
```

**Returns:**
- `tables`: List of table info (column widths, row heights, cell properties)
- `paragraphs_count`: Number of paragraphs
- `has_header`: Whether document has header
- `has_footer`: Whether document has footer

#### Compare Documents (NEW)

Compares the structure of two Word documents.

```python
compare_word_documents(filepath1: str, filepath2: str) -> Dict
```

**Returns:**
- `is_identical`: Boolean indicating structural match
- `differences`: List of structural differences found

#### Create a Word Document

```python
create_word_document(filepath: str, content: str) -> Dict
```

Creates a simple Word document with plain text content.

#### Edit a Word Document

```python
edit_word_document(filepath: str, operations: List[Dict]) -> Dict
```

**Supported operations:**
- `add_paragraph`: Add a new paragraph
- `add_heading`: Add a heading (level 1-4)
- `edit_paragraph`: Modify existing paragraph by index
- `delete_paragraph`: Remove paragraph by index

#### Convert TXT to Word

```python
convert_txt_to_word(source_path: str, target_path: str) -> Dict
```

### Excel

#### Create an Excel File

```python
create_excel_file(filepath: str, content: str) -> Dict
```

#### Edit an Excel File

```python
edit_excel_file(filepath: str, operations: List[Dict]) -> Dict
```

**Supported operations:**
- `update_cell`: Update a single cell
- `update_range`: Update a range of cells
- `delete_row`: Delete a row
- `delete_column`: Delete a column
- `add_sheet`: Add a new sheet
- `delete_sheet`: Delete a sheet

#### Convert CSV to Excel

```python
convert_csv_to_excel(source_path: str, target_path: str) -> Dict
```

### PDF

#### Create a PDF File

```python
create_pdf_file(filepath: str, content: str) -> Dict
```

#### Convert Word to PDF

```python
convert_word_to_pdf(source_path: str, target_path: str) -> Dict
```

## Logs

The server logs all operations to both the console and a `logs/document_mcp.log` file for troubleshooting.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Changelog

### v0.2.0 (2024-12)
- Added `create_formatted_word_document` for rich document creation
- Added `read_word_document_structure` for document analysis
- Added `compare_word_documents` for structural comparison
- Added support for tables with custom styling
- Added support for headers and footers
- Added rich text parsing (`**bold**` syntax)
- Added key-value table type for forms

### v0.1.0
- Initial release with basic Word, Excel, and PDF operations
