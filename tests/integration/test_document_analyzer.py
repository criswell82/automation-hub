"""
Test suite for DocumentAnalyzer - Phase 2 verification.

Creates test documents and verifies analysis capabilities.
"""
import sys
import json
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.modules.word_automation.document_handler import DocumentHandler
from src.core.document_analyzer import DocumentAnalyzer, TemplateMode


def create_fill_in_template():
    """Create a document with placeholders (FILL_IN mode)."""
    print("Creating FILL_IN test document...")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Customer Invoice", level=1)
    handler.add_paragraph("")
    handler.add_paragraph("Bill To: {{client_name}}")
    handler.add_paragraph("Email: {{client_email}}")
    handler.add_paragraph("Invoice Date: {{invoice_date}}")
    handler.add_paragraph("")

    handler.add_heading("Invoice Details", level=2)
    handler.add_table(
        data=[
            ['{{item1_description}}', '{{item1_qty}}', '{{item1_price}}'],
            ['{{item2_description}}', '{{item2_qty}}', '{{item2_price}}'],
        ],
        headers=['Description', 'Quantity', 'Unit Price']
    )

    handler.add_paragraph("")
    handler.add_paragraph("Total Amount: {{total_amount}}")
    handler.add_paragraph("")
    handler.add_paragraph("Payment Due: {{due_date}}")

    output_path = project_root / "test_output" / "fill_in_template.docx"
    handler.save(str(output_path))

    print(f"  [OK] Created: {output_path.name}")
    return str(output_path)


def create_generate_template():
    """Create a complex document (GENERATE mode)."""
    print("Creating GENERATE test document...")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Quarterly Business Report", level=1)
    handler.add_paragraph("Executive Summary:")
    handler.add_paragraph("This report provides an overview of business performance for Q3 2025.")

    handler.add_heading("Revenue Analysis", level=2)
    handler.add_paragraph("Total revenue reached $500,000 this quarter.")
    handler.add_table(
        data=[
            ['Product A', '$200,000'],
            ['Product B', '$150,000'],
            ['Product C', '$150,000']
        ],
        headers=['Product', 'Revenue']
    )

    handler.add_heading("Expenses Breakdown", level=2)
    handler.add_table(
        data=[
            ['Salaries', '$150,000'],
            ['Marketing', '$75,000'],
            ['Operations', '$100,000']
        ],
        headers=['Category', 'Amount']
    )

    handler.add_heading("Conclusions", level=2)
    handler.add_paragraph("The quarter showed strong growth with opportunities for expansion.")

    output_path = project_root / "test_output" / "generate_template.docx"
    handler.save(str(output_path))

    print(f"  [OK] Created: {output_path.name}")
    return str(output_path)


def create_content_template():
    """Create a simple static document (CONTENT mode)."""
    print("Creating CONTENT test document...")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Company Letterhead", level=1)
    handler.add_paragraph("")
    handler.add_paragraph("Acme Corporation")
    handler.add_paragraph("123 Business Street")
    handler.add_paragraph("City, State 12345")
    handler.add_paragraph("")
    handler.add_paragraph("This is a standard company letterhead template.")

    output_path = project_root / "test_output" / "content_template.docx"
    handler.save(str(output_path))

    print(f"  [OK] Created: {output_path.name}")
    return str(output_path)


def create_pattern_template():
    """Create a document with repeating structure (PATTERN mode)."""
    print("Creating PATTERN test document...")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Employee Directory", level=1)

    # Multiple similar tables (repeating pattern)
    handler.add_heading("Engineering Department", level=2)
    handler.add_table(
        data=[['John Doe', 'Engineer'], ['Jane Smith', 'Senior Engineer']],
        headers=['Name', 'Title']
    )

    handler.add_heading("Sales Department", level=2)
    handler.add_table(
        data=[['Bob Wilson', 'Sales Rep'], ['Alice Brown', 'Sales Manager']],
        headers=['Name', 'Title']
    )

    handler.add_heading("Marketing Department", level=2)
    handler.add_table(
        data=[['Carol White', 'Marketing Specialist'], ['Dan Green', 'Marketing Director']],
        headers=['Name', 'Title']
    )

    output_path = project_root / "test_output" / "pattern_template.docx"
    handler.save(str(output_path))

    print(f"  [OK] Created: {output_path.name}")
    return str(output_path)


def test_analyzer(filepath, expected_mode, test_name):
    """Test analyzer on a document."""
    print(f"\n{'=' * 60}")
    print(f"TEST: {test_name}")
    print(f"{'=' * 60}")

    analyzer = DocumentAnalyzer()
    analysis = analyzer.analyze_word_document(filepath)

    print(f"\nDetected Mode: {analysis['mode']}")
    print(f"Confidence: {analysis['confidence']:.2f}")
    print(f"Expected Mode: {expected_mode}")

    # Verify mode
    if analysis['mode'] == expected_mode:
        print(f"[OK] Mode detection correct")
    else:
        print(f"[WARN] Mode mismatch - expected {expected_mode}, got {analysis['mode']}")

    # Display structure info
    structure = analysis['structure']
    print(f"\nStructure:")
    print(f"  Paragraphs: {structure.get('total_paragraphs', 0)}")
    print(f"  Tables: {structure.get('total_tables', 0)}")
    print(f"  Headings: {len(structure.get('headings', []))}")
    print(f"  Complexity: {structure.get('complexity_score', 0)}")

    # Display placeholders
    placeholders = analysis['placeholders']
    if placeholders:
        print(f"\nPlaceholders Found: {len(placeholders)}")
        for p in placeholders[:5]:  # Show first 5
            print(f"  - {p['name']} ({p['type']}): {p['description']}")
            print(f"    Pattern: {p['pattern']}, Count: {p['count']}")
        if len(placeholders) > 5:
            print(f"  ... and {len(placeholders) - 5} more")
    else:
        print("\nPlaceholders Found: 0")

    # Display parameters
    parameters = analysis['parameters']
    if parameters:
        print(f"\nGenerated Parameters: {len(parameters)}")
        for param_name, param_def in list(parameters.items())[:5]:
            print(f"  - {param_name}:")
            print(f"      type: {param_def['type']}")
            print(f"      description: {param_def['description']}")
            print(f"      required: {param_def.get('required', False)}")
        if len(parameters) > 5:
            print(f"  ... and {len(parameters) - 5} more")

    # Display recommendations
    print(f"\nRecommendations:")
    print(f"  Template Name: {analysis['recommended_template_name']}")
    print(f"  Category: {analysis['recommended_category']}")

    return analysis


def test_placeholder_detection():
    """Test various placeholder patterns."""
    print(f"\n{'=' * 60}")
    print(f"TEST: Placeholder Pattern Detection")
    print(f"{'=' * 60}")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Placeholder Pattern Tests", level=1)
    handler.add_paragraph("Jinja2 style: {{customer_name}}")
    handler.add_paragraph("Bracket style: [CUSTOMER_EMAIL]")
    handler.add_paragraph("Brace style: {customer_phone}")
    handler.add_paragraph("Angle style: <customer_address>")
    handler.add_paragraph("Underscore style: __customer_id__")

    output_path = project_root / "test_output" / "placeholder_patterns.docx"
    handler.save(str(output_path))

    analyzer = DocumentAnalyzer()
    analysis = analyzer.analyze_word_document(str(output_path))

    placeholders = analysis['placeholders']
    print(f"\nDetected {len(placeholders)} placeholders:")
    for p in placeholders:
        print(f"  - {p['name']} with pattern '{p['pattern']}'")

    expected_count = 5
    if len(placeholders) == expected_count:
        print(f"\n[OK] All {expected_count} placeholder patterns detected")
        return True
    else:
        print(f"\n[WARN] Expected {expected_count} placeholders, found {len(placeholders)}")
        return False


def test_type_inference():
    """Test parameter type inference."""
    print(f"\n{'=' * 60}")
    print(f"TEST: Type Inference")
    print(f"{'=' * 60}")

    handler = DocumentHandler()
    handler.create_document()

    handler.add_heading("Type Inference Tests", level=1)
    handler.add_paragraph("Date field: {{invoice_date}}")
    handler.add_paragraph("Email field: {{client_email}}")
    handler.add_paragraph("Number field: {{total_amount}}")
    handler.add_paragraph("File field: {{attachment_file}}")
    handler.add_paragraph("Boolean field: {{is_active}}")
    handler.add_paragraph("String field: {{customer_name}}")

    output_path = project_root / "test_output" / "type_inference.docx"
    handler.save(str(output_path))

    analyzer = DocumentAnalyzer()
    analysis = analyzer.analyze_word_document(str(output_path))

    expected_types = {
        'invoice_date': 'date',
        'client_email': 'email',
        'total_amount': 'number',
        'attachment_file': 'file',
        'is_active': 'boolean',
        'customer_name': 'string'
    }

    print("\nType Inference Results:")
    all_correct = True
    for placeholder in analysis['placeholders']:
        name = placeholder['name']
        detected_type = placeholder['type']
        expected_type = expected_types.get(name, 'unknown')

        status = "[OK]" if detected_type == expected_type else "[WARN]"
        print(f"  {status} {name}: {detected_type} (expected: {expected_type})")

        if detected_type != expected_type:
            all_correct = False

    if all_correct:
        print("\n[OK] All types inferred correctly")
    else:
        print("\n[WARN] Some type inferences incorrect")

    return all_correct


def save_analysis_report(analyses):
    """Save analysis results to JSON for inspection."""
    report_path = project_root / "test_output" / "analysis_report.json"

    report = {}
    for test_name, analysis in analyses.items():
        # Simplify for JSON
        report[test_name] = {
            'mode': analysis['mode'],
            'confidence': analysis['confidence'],
            'placeholder_count': len(analysis['placeholders']),
            'parameter_count': len(analysis['parameters']),
            'template_name': analysis['recommended_template_name'],
            'category': analysis['recommended_category'],
            'complexity': analysis['structure'].get('complexity_score', 0)
        }

    with open(report_path, 'w') as f:
        json.dump(report, f, indent=2)

    print(f"\n[OK] Analysis report saved: {report_path.name}")


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("DOCUMENT ANALYZER TEST SUITE - PHASE 2")
    print("=" * 60)
    print()

    try:
        # Create test documents
        print("Step 1: Creating test documents")
        print("-" * 60)
        fill_in_path = create_fill_in_template()
        generate_path = create_generate_template()
        content_path = create_content_template()
        pattern_path = create_pattern_template()

        # Test analyzer
        print("\n\nStep 2: Testing analyzer on documents")
        print("-" * 60)

        analyses = {}
        analyses['fill_in'] = test_analyzer(fill_in_path, 'fill_in', "FILL_IN Mode Detection")
        analyses['generate'] = test_analyzer(generate_path, 'generate', "GENERATE Mode Detection")
        analyses['content'] = test_analyzer(content_path, 'content', "CONTENT Mode Detection")
        analyses['pattern'] = test_analyzer(pattern_path, 'pattern', "PATTERN Mode Detection")

        # Test specific features
        print("\n\nStep 3: Testing specific features")
        print("-" * 60)
        placeholder_test = test_placeholder_detection()
        type_test = test_type_inference()

        # Save report
        save_analysis_report(analyses)

        # Summary
        print("\n\n" + "=" * 60)
        print("TEST SUMMARY")
        print("=" * 60)
        print("\nDocuments Created:")
        print("  [OK] FILL_IN template")
        print("  [OK] GENERATE template")
        print("  [OK] CONTENT template")
        print("  [OK] PATTERN template")

        print("\nAnalyzer Tests:")
        for test_name, analysis in analyses.items():
            status = "[OK]" if analysis['confidence'] > 0.7 else "[WARN]"
            print(f"  {status} {test_name.upper()} mode - confidence: {analysis['confidence']:.2f}")

        print("\nFeature Tests:")
        print(f"  {'[OK]' if placeholder_test else '[WARN]'} Placeholder detection")
        print(f"  {'[OK]' if type_test else '[WARN]'} Type inference")

        print("\n" + "=" * 60)
        print("PHASE 2 TESTS COMPLETED")
        print("=" * 60)
        print()

    except Exception as e:
        print(f"\n[FAIL] TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
