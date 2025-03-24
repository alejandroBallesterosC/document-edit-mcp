#!/usr/bin/env python3
"""
Simple test client for the Document MCP Server
"""

import asyncio
import json
import os
import sys
import tempfile
from pathlib import Path
from contextlib import asynccontextmanager

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client


async def test_word_operations(session):
    print("\n\nTesting Word operations...")
    
    # Create a temp directory for test files
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create a text file
        txt_path = os.path.join(temp_dir, "test.txt")
        with open(txt_path, "w") as f:
            f.write("This is a test text file.\nIt has multiple lines.\nWe'll convert it to Word.")
        
        # Create a Word document
        word_path = os.path.join(temp_dir, "test_created.docx")
        
        print("Creating Word document...")
        result = await session.call_tool(
            "create_word_document",
            {"filepath": word_path, "content": "This is a test Word document created by the MCP server."}
        )
        print(f"Result: {json.dumps(result, indent=2)}")
        
        if result.get("success"):
            # Edit the Word document
            print("\nEditing Word document...")
            result = await session.call_tool(
                "edit_word_document",
                {
                    "filepath": word_path,
                    "operations": [
                        {"type": "add_paragraph", "text": "This is a new paragraph."},
                        {"type": "add_heading", "text": "Test Heading", "level": 1}
                    ]
                }
            )
            print(f"Result: {json.dumps(result, indent=2)}")
        
        # Convert text to Word
        word_converted_path = os.path.join(temp_dir, "test_converted.docx")
        print("\nConverting TXT to Word...")
        result = await session.call_tool(
            "convert_txt_to_word",
            {"source_path": txt_path, "target_path": word_converted_path}
        )
        print(f"Result: {json.dumps(result, indent=2)}")


async def test_capabilities(session):
    print("\n\nGetting server capabilities...")
    result = await session.get_resource("capabilities://")
    print(f"Capabilities: {json.dumps(json.loads(result), indent=2)}")


@asynccontextmanager
async def connect_to_server():
    """Helper to connect to the MCP server."""
    server_params = StdioServerParameters(
        command="python",
        args=["-m", "claude_document_mcp.server"],
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    
    client = stdio_client(server_params)
    session = await client.connect()
    
    try:
        yield session
    finally:
        await session.close()


async def main():
    print("Document MCP Server Test Client")
    print("===============================")
    
    try:
        # Connect to the MCP server using stdio
        print("Connecting to server...")
        
        async with connect_to_server() as session:
            print("Connected!")
            
            # Get server capabilities
            await test_capabilities(session)
            
            # Test Word operations
            await test_word_operations(session)
            
            print("\nAll tests completed!")
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())
