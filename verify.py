#!/usr/bin/env python3
"""
Verification script for Claude Document MCP Server

This script verifies that the environment is correctly set up and that
all dependencies are installed properly.
"""

import sys
import importlib
import os
import subprocess
from pathlib import Path

def run_uv_command(args):
    """Run a UV command and return the output."""
    cmd = ["uv"] + args
    try:
        return subprocess.check_output(cmd, universal_newlines=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running UV command: {e}")
        return None

def check_dependency(module_name):
    """Check if a Python module is installed."""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False

def main():
    """Main verification function."""
    print("Claude Document MCP Server Verification")
    print("======================================")
    
    # Check if UV is installed
    try:
        uv_version = subprocess.check_output(["uv", "--version"], universal_newlines=True).strip()
        print(f"UV installed: Yes (version {uv_version})")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("UV installed: No")
        print("ERROR: UV is not installed or not in PATH. Please install UV first.")
        return False
    
    # Check Python version
    print(f"Python version: {sys.version}")
    
    # Check if the project is installed
    print("\nChecking dependencies...")
    
    # Try to import the project
    can_import = check_dependency("claude_document_mcp")
    print(f"claude_document_mcp importable: {'Yes' if can_import else 'No'}")
    
    # Check MCP dependency
    if check_dependency("mcp"):
        import mcp
        print(f"MCP installed: Yes (version {getattr(mcp, '__version__', 'unknown')})")
    else:
        print("MCP installed: No")
    
    # Check other dependencies
    dependencies = [
        "docx", 
        "pandas", 
        "openpyxl", 
        "reportlab", 
        "pdf2docx", 
        "docx2pdf"
    ]
    
    missing_deps = []
    for dep in dependencies:
        installed = check_dependency(dep)
        print(f"{dep} installed: {'Yes' if installed else 'No'}")
        if not installed:
            missing_deps.append(dep)
    
    if missing_deps:
        print(f"\nMissing dependencies: {', '.join(missing_deps)}")
        print("Run 'uv sync' to install missing dependencies")
    
    # Check if logs directory exists
    logs_dir = Path(__file__).parent / "logs"
    if logs_dir.exists():
        print(f"\nLogs directory exists: Yes ({logs_dir})")
    else:
        print(f"\nLogs directory exists: No")
        print(f"Creating logs directory at {logs_dir}")
        logs_dir.mkdir(exist_ok=True)
    
    # Check if Claude Desktop config exists
    if sys.platform == "darwin":
        config_path = Path.home() / "Library/Application Support/Claude/claude_desktop_config.json"
    elif sys.platform == "win32":
        config_path = Path(os.environ.get("APPDATA", "")) / "Claude/claude_desktop_config.json"
    else:
        config_path = Path(__file__).parent / "claude_desktop_config.json"
    
    if config_path.exists():
        print(f"Claude Desktop config exists: Yes ({config_path})")
    else:
        print(f"Claude Desktop config exists: No")
        print(f"ERROR: Claude Desktop config does not exist at {config_path}")
        print("Run './setup.sh' to create the configuration")
        return False
    
    # Test running the server with UV
    print("\nTesting MCP server execution with UV...")
    try:
        # Just check if the command would run, don't actually run it
        cmd = ["uv", "run", "-m", "claude_document_mcp.server", "--help"]
        subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("Server command is executable: Yes")
    except subprocess.CalledProcessError:
        print("Server command is executable: No")
        print("ERROR: Cannot execute the MCP server")
        return False
    
    print("\nVerification successful! Your environment is properly set up.")
    print("To run the server, use: ./run.sh")
    print("Or directly with UV: uv run -m claude_document_mcp.server")
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
