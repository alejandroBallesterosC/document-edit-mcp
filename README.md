[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/alejandroballesterosc-document-edit-mcp-badge.png)](https://mseep.ai/app/alejandroballesterosc-document-edit-mcp)

# Claude Document MCP Server
[![smithery badge](https://smithery.ai/badge/@alejandroBallesterosC/document-edit-mcp)](https://smithery.ai/server/@alejandroBallesterosC/document-edit-mcp)

A Model Context Protocol (MCP) server that allows Claude Desktop to perform document operations on Microsoft Word, Excel, and PDF files.

## Features

### Microsoft Word Operations
- Create new Word documents from text
- Edit existing Word documents (add/edit/delete paragraphs and headings)
- Convert text files (.txt) to Word documents

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

#### Create a Word Document
```
create_word_document(filepath: str, content: str) -> Dict
```

#### Edit a Word Document
```
edit_word_document(filepath: str, operations: List[Dict]) -> Dict
```

#### Convert TXT to Word
```
convert_txt_to_word(source_path: str, target_path: str) -> Dict
```

### Excel

#### Create an Excel File
```
create_excel_file(filepath: str, content: str) -> Dict
```

#### Edit an Excel File
```
edit_excel_file(filepath: str, operations: List[Dict]) -> Dict
```

#### Convert CSV to Excel
```
convert_csv_to_excel(source_path: str, target_path: str) -> Dict
```

### PDF

#### Create a PDF File
```
create_pdf_file(filepath: str, content: str) -> Dict
```

#### Convert Word to PDF
```
convert_word_to_pdf(source_path: str, target_path: str) -> Dict
```

## Logs

The server logs all operations to both the console and a `logs/document_mcp.log` file for troubleshooting.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
