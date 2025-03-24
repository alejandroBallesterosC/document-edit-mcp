#!/bin/bash
# Setup script for Claude Document MCP Server

set -e  # Exit on error

echo "Setting up Claude Document MCP Server..."

# Check Python version
python_version=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
min_version="3.10"

if [ "$(printf '%s\n' "$min_version" "$python_version" | sort -V | head -n1)" != "$min_version" ]; then 
    echo "Error: Python $min_version or higher is required"
    echo "Current version: $python_version"
    exit 1
fi

# Create virtual environment first (using UV)
echo "Creating virtual environment with UV..."
uv sync

# Now install the project in development mode
echo "Installing project in development mode with UV..."
uv pip install -e .

# Create logs directory
mkdir -p logs

# Get the absolute path to the project directory
PROJECT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)

echo ""
echo "Setup complete! You can now use the Document MCP Server with Claude Desktop."
echo "Start Claude Desktop to use the document tools."
echo ""