"""
Tests for the new formatted document functions.

Run with: pytest tests/test_formatted_documents.py -v
"""

import json
import os
import tempfile
import pytest
from pathlib import Path

# Add parent directory to path for imports
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from claude_document_mcp.server import (
    create_formatted_word_document,
    read_word_document_structure,
    compare_word_documents,
    create_word_document
)


class TestCreateFormattedWordDocument:
    """Tests for create_formatted_word_document function."""

    def test_basic_document_creation(self, tmp_path):
        """Test creating a simple formatted document."""
        filepath = str(tmp_path / "test_basic.docx")
        
        document_data = json.dumps({
            "title": "Test Document",
            "sections": [
                {"type": "paragraph", "text": "Hello World"}
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_document_with_heading(self, tmp_path):
        """Test creating document with headings."""
        filepath = str(tmp_path / "test_heading.docx")
        
        document_data = json.dumps({
            "sections": [
                {"type": "heading", "level": 1, "text": "Main Title"},
                {"type": "heading", "level": 2, "text": "Subtitle"},
                {"type": "paragraph", "text": "Content here."}
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_document_with_table(self, tmp_path):
        """Test creating document with a data table."""
        filepath = str(tmp_path / "test_table.docx")
        
        document_data = json.dumps({
            "sections": [
                {
                    "type": "table",
                    "headers": ["Name", "Value"],
                    "rows": [
                        ["Item 1", "100"],
                        ["Item 2", "200"]
                    ],
                    "header_bg_color": "1F4E79",
                    "header_text_color": "FFFFFF"
                }
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_document_with_key_value_table(self, tmp_path):
        """Test creating document with a key-value table."""
        filepath = str(tmp_path / "test_kv_table.docx")
        
        document_data = json.dumps({
            "sections": [
                {
                    "type": "key_value_table",
                    "rows": [
                        ["Project", "Alpha"],
                        ["Status", "Active"],
                        ["Owner", "John"]
                    ]
                }
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_document_with_lists(self, tmp_path):
        """Test creating document with bullet and numbered lists."""
        filepath = str(tmp_path / "test_lists.docx")
        
        document_data = json.dumps({
            "sections": [
                {
                    "type": "bullet_list",
                    "items": ["First item", "Second item", "Third item"]
                },
                {
                    "type": "numbered_list",
                    "items": ["Step 1", "Step 2", "Step 3"]
                }
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_document_with_header_footer(self, tmp_path):
        """Test creating document with header and footer."""
        filepath = str(tmp_path / "test_header_footer.docx")
        
        document_data = json.dumps({
            "title": "Report",
            "header": "Confidential - Internal Use Only",
            "footer": "© 2024 Company Name",
            "sections": [
                {"type": "paragraph", "text": "Document content."}
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_rich_text_bold(self, tmp_path):
        """Test rich text parsing with bold markers."""
        filepath = str(tmp_path / "test_bold.docx")
        
        document_data = json.dumps({
            "sections": [
                {
                    "type": "paragraph",
                    "text": "This has **bold text** in the middle."
                }
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)

    def test_invalid_json(self, tmp_path):
        """Test handling of invalid JSON input."""
        filepath = str(tmp_path / "test_invalid.docx")
        
        result = create_formatted_word_document(filepath, "not valid json")
        
        assert result["success"] is False
        assert "Invalid JSON" in result["message"]

    def test_full_document(self, tmp_path):
        """Test creating a complete document with all features."""
        filepath = str(tmp_path / "test_full.docx")
        
        document_data = json.dumps({
            "title": "Complete Test Document",
            "subtitle": "Testing All Features",
            "header": "Test Header",
            "footer": "Test Footer - Page 1",
            "sections": [
                {"type": "heading", "level": 1, "text": "Introduction", "color": "1F4E79"},
                {"type": "paragraph", "text": "This is a **comprehensive** test.", "alignment": "justify"},
                {"type": "heading", "level": 2, "text": "Data Table"},
                {
                    "type": "table",
                    "headers": ["A", "B", "C"],
                    "rows": [["1", "2", "3"], ["4", "5", "6"]],
                    "header_bg_color": "2E75B6"
                },
                {"type": "heading", "level": 2, "text": "Key Information"},
                {
                    "type": "key_value_table",
                    "rows": [["Key", "Value"], ["Name", "Test"]]
                },
                {"type": "bullet_list", "items": ["Point A", "Point B"]},
                {"type": "page_break"},
                {"type": "heading", "level": 1, "text": "Page 2"},
                {"type": "paragraph", "text": "Content on second page."}
            ]
        })
        
        result = create_formatted_word_document(filepath, document_data)
        
        assert result["success"] is True
        assert os.path.exists(filepath)


class TestReadWordDocumentStructure:
    """Tests for read_word_document_structure function."""

    def test_read_simple_document(self, tmp_path):
        """Test reading structure of a simple document."""
        # First create a document
        filepath = str(tmp_path / "test_read.docx")
        create_word_document(filepath, "Simple content")
        
        result = read_word_document_structure(filepath)
        
        assert result["success"] is True
        assert "paragraphs_count" in result
        assert result["paragraphs_count"] >= 1

    def test_read_document_with_table(self, tmp_path):
        """Test reading structure of document with table."""
        filepath = str(tmp_path / "test_read_table.docx")
        
        document_data = json.dumps({
            "sections": [
                {
                    "type": "table",
                    "headers": ["A", "B"],
                    "rows": [["1", "2"], ["3", "4"]]
                }
            ]
        })
        create_formatted_word_document(filepath, document_data)
        
        result = read_word_document_structure(filepath)
        
        assert result["success"] is True
        assert result["tables_count"] == 1
        assert len(result["tables"]) == 1

    def test_read_nonexistent_file(self, tmp_path):
        """Test handling of non-existent file."""
        filepath = str(tmp_path / "nonexistent.docx")
        
        result = read_word_document_structure(filepath)
        
        assert result["success"] is False


class TestCompareWordDocuments:
    """Tests for compare_word_documents function."""

    def test_compare_identical_documents(self, tmp_path):
        """Test comparing two identical documents."""
        filepath1 = str(tmp_path / "doc1.docx")
        filepath2 = str(tmp_path / "doc2.docx")
        
        document_data = json.dumps({
            "sections": [
                {"type": "paragraph", "text": "Same content"}
            ]
        })
        
        create_formatted_word_document(filepath1, document_data)
        create_formatted_word_document(filepath2, document_data)
        
        result = compare_word_documents(filepath1, filepath2)
        
        assert result["success"] is True
        assert result["is_identical"] is True

    def test_compare_different_documents(self, tmp_path):
        """Test comparing two different documents."""
        filepath1 = str(tmp_path / "doc1.docx")
        filepath2 = str(tmp_path / "doc2.docx")
        
        # Document with table
        doc1_data = json.dumps({
            "sections": [
                {
                    "type": "table",
                    "headers": ["A", "B"],
                    "rows": [["1", "2"]]
                }
            ]
        })
        
        # Document without table
        doc2_data = json.dumps({
            "sections": [
                {"type": "paragraph", "text": "No table here"}
            ]
        })
        
        create_formatted_word_document(filepath1, doc1_data)
        create_formatted_word_document(filepath2, doc2_data)
        
        result = compare_word_documents(filepath1, filepath2)
        
        assert result["success"] is True
        assert result["is_identical"] is False
        assert len(result["differences"]) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
