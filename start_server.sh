#!/bin/bash

# Start the Claude Document MCP Server

# Activate the virtual environment
if [ -d ".venv" ]; then
    source .venv/bin/activate
else
    echo "Creating virtual environment with UV..."
    uv venv
    source .venv/bin/activate
    
    echo "Installing dependencies with UV..."
    uv pip install -e .
fi

# Start the server
echo "Starting Document MCP Server..."
python -m claude_document_mcp
