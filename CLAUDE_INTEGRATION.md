# Integrating with Claude Desktop

To integrate this Document MCP Server with Claude Desktop, follow these steps:

## 1. Install the Server

First, make sure you have installed the server with dependencies:

```bash
cd claude-document-mcp
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install -e .
```

## 2. Configure Claude Desktop

1. Find your Claude Desktop configuration file:
   - **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. Create or edit this file to include the following configuration:

```json
{
  "mcpServers": {
    "document_operations": {
      "command": "uv",
      "args": [
        "--directory", "/ABSOLUTE/PATH/TO/claude-document-mcp", 
        "run", 
        "mcp", 
        "run", 
        "claude_document_mcp/server.py:mcp"
      ]
    }
  }
}
```

Replace `/ABSOLUTE/PATH/TO/claude-document-mcp` with the actual absolute path to your project directory.

3. Restart Claude Desktop

## 3. Test the Integration

1. Open Claude Desktop
2. You should see a hammer icon in the UI if the MCP server is detected
3. Try a test request like: "Can you create a new Word document with a simple hello world message and save it to my Desktop?"

## Troubleshooting

If Claude Desktop doesn't detect the server:

1. Check the logs directory for errors
2. Verify your Claude Desktop config file has the correct path
3. Make sure the MCP server runs correctly on its own (use `./test_server.sh`)
4. Restart Claude Desktop after any changes
