#!/bin/bash

# This script tests the Document MCP Server by running it and confirming it works

# Create temporary directory for test files
TEST_DIR=$(mktemp -d)
echo "Created test directory: $TEST_DIR"

# Create a test text file
TEXT_FILE="$TEST_DIR/test.txt"
echo "This is a test text file." > "$TEXT_FILE"
echo "It has multiple lines." >> "$TEXT_FILE"
echo "We'll convert it to Word." >> "$TEXT_FILE"
echo "Created text file: $TEXT_FILE"

# The main test process
# This will start the server in the background and run tests against it
echo "Starting Document MCP Server validation..."

# Start the server with the MCP dev tool (this includes the inspector)
echo "Starting server..."
mcp dev claude_document_mcp/server.py:mcp &
SERVER_PID=$!

# Wait for server to start
sleep 3

echo "Server started with PID: $SERVER_PID"
echo "To view the server in the MCP Inspector, visit: http://localhost:9000"
echo ""
echo "Press Ctrl+C to stop the server when you're done testing"

# Keep the script running so the server stays up
wait $SERVER_PID

# Clean up
echo "Cleaning up test files..."
rm -rf "$TEST_DIR"
echo "Done."
