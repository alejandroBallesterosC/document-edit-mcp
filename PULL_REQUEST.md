# Pull Request: File Management with Trash Support (v0.3.1)

## Summary

This PR adds file management capabilities to the document-edit-mcp server, including the ability to delete files and directories with two safety modes:
- **Trash mode**: Files are sent to the system's Recycle Bin (Windows) / Trash (macOS/Linux) and can be recovered
- **Permanent mode**: Files are permanently deleted and cannot be recovered

## Motivation

When working with Claude Desktop to manage documents, users often need to clean up old files or temporary documents. Previously, this required manual intervention. This PR adds safe file deletion with explicit user confirmation to prevent accidental data loss.

## Changes

### New Functions

#### `delete_file(filepath, confirm)`
Delete a file with explicit user confirmation.
- `confirm="CORBEILLE"`: Send to system trash (recoverable)
- `confirm="SUPPRESSION DÉFINITIVE"`: Permanent deletion

#### `delete_directory(dirpath, confirm)`
Delete an empty directory with the same safety modes.
- Returns an error with contents list if directory is not empty

### New Dependency
- `send2trash>=1.8.0` - Cross-platform trash/recycle bin support

### Safety Features
- **Explicit confirmation required**: No accidental deletions
- **Two-mode system**: Users choose between recoverable and permanent
- **Directory protection**: Cannot delete non-empty directories
- **Detailed feedback**: Returns file size, path, and deletion method

## Testing

```python
# Test trash mode
delete_file("test.docx", "CORBEILLE")
# -> File appears in Recycle Bin, can be restored

# Test permanent mode  
delete_file("test.docx", "SUPPRESSION DÉFINITIVE")
# -> File permanently deleted

# Test directory protection
delete_directory("non_empty_folder", "CORBEILLE")
# -> Error: "Directory not empty (X items). Delete contents first."
```

## Backward Compatibility

- No changes to existing functions
- New functions are additive only
- Existing workflows continue to work unchanged

## Installation Note

After updating, run:
```bash
pip install send2trash
```

Or use the provided script:
```bash
scripts/install_dependencies.bat
```

## Checklist

- [x] Code follows project style guidelines
- [x] Functions documented with docstrings
- [x] CHANGELOG.md updated
- [x] README.md updated with new API documentation
- [x] pyproject.toml updated with new dependency
- [x] Installation scripts updated
