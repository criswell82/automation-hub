"""
Test script for DocumentHandler to verify Phase 1 implementation.

Creates a sample document, tests placeholder replacement, and structure extraction.
"""
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.modules.word_automation.document_handler import DocumentHandler

def test_document_creation():
    """Test creating a new document with content."""
    print("=" * 60)
    print("TEST 1: Document Creation")
    print("=" * 60)

    handler = DocumentHandler()
    handler.create_document()

    # Add content
    handler.add_heading("Sample Invoice Template", level=1)
    handler.add_paragraph("Thank you for your business!")
    handler.add_paragraph("")
    handler.add_heading("Client Information", level=2)
    handler.add_paragraph("Client: {{client_name}}")
    handler.add_paragraph("Date: {{invoice_date}}")
    handler.add_paragraph("Amount: {{amount}}")

    # Add a table
    handler.add_table(
        data=[
            ['Item 1', '{{item1_qty}}', '{{item1_price}}'],
            ['Item 2', '{{item2_qty}}', '{{item2_price}}'],
        ],
        headers=['Description', 'Quantity', 'Price']
    )

    # Save
    output_path = project_root / "test_output" / "sample_template.docx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    handler.save(str(output_path))

    print(f"[OK] Created document: {output_path}")
    print()

    return str(output_path)


def test_structure_extraction(filepath):
    """Test extracting structure from document."""
    print("=" * 60)
    print("TEST 2: Structure Extraction")
    print("=" * 60)

    handler = DocumentHandler(filepath)
    structure = handler.extract_structure()

    print(f"Total paragraphs: {structure['total_paragraphs']}")
    print(f"Total tables: {structure['total_tables']}")
    print(f"\nHeadings found:")
    for heading in structure['headings']:
        print(f"  Level {heading['level']}: {heading['text']}")

    print(f"\nPlaceholders found: {', '.join(structure['placeholders'])}")
    print(f"\nTable structure:")
    for table in structure['tables']:
        print(f"  Table {table['index']}: {table['rows']} rows x {table['cols']} cols")

    print("[OK] Structure extracted successfully")
    print()

    return structure


def test_placeholder_replacement(filepath):
    """Test replacing placeholders."""
    print("=" * 60)
    print("TEST 3: Placeholder Replacement")
    print("=" * 60)

    handler = DocumentHandler(filepath)

    # Replace placeholders
    handler.replace_placeholders({
        'client_name': 'Acme Corporation',
        'invoice_date': '2025-11-04',
        'amount': '$2,500.00',
        'item1_qty': '10',
        'item1_price': '$100.00',
        'item2_qty': '15',
        'item2_price': '$100.00'
    })

    # Save filled document
    output_path = project_root / "test_output" / "filled_invoice.docx"
    handler.save(str(output_path))

    print(f"[OK] Placeholders replaced and saved: {output_path}")
    print()


def test_document_builder():
    """Test building a document from scratch."""
    print("=" * 60)
    print("TEST 4: Building Document from Scratch")
    print("=" * 60)

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Monthly Report", level=1)
    handler.add_paragraph("This is an automated monthly report.")

    handler.add_heading("Summary", level=2)
    handler.add_paragraph("Total sales: $10,000")
    handler.add_paragraph("Total expenses: $5,000")
    handler.add_paragraph("Net profit: $5,000")

    handler.add_heading("Details", level=2)
    handler.add_table(
        data=[
            ['Q1', '$2,500'],
            ['Q2', '$3,000'],
            ['Q3', '$2,500'],
            ['Q4', '$2,000']
        ],
        headers=['Quarter', 'Sales']
    )

    output_path = project_root / "test_output" / "monthly_report.docx"
    handler.save(str(output_path))

    print(f"[OK] Built document from scratch: {output_path}")
    print()


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("DOCUMENT HANDLER TEST SUITE")
    print("=" * 60)
    print()

    try:
        # Run tests
        template_path = test_document_creation()
        structure = test_structure_extraction(template_path)
        test_placeholder_replacement(template_path)
        test_document_builder()

        print("=" * 60)
        print("ALL TESTS PASSED [OK]")
        print("=" * 60)
        print()
        print("Output files created in test_output/:")
        print("  - sample_template.docx (template with placeholders)")
        print("  - filled_invoice.docx (template with placeholders filled)")
        print("  - monthly_report.docx (document built from scratch)")
        print()

    except Exception as e:
        print(f"\n[FAIL] TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
