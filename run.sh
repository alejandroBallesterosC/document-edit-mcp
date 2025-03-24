#!/bin/bash
# Run the Claude Document MCP Server directly

# Get the project directory
PROJECT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)

# Run the server with UV (using --project parameter)
echo "Starting Document MCP Server..."
uv run --project "$PROJECT_DIR" -m claude_document_mcp.server