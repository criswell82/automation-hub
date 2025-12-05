"""
Test OneNote Integration
Validates that all OneNote components are properly integrated.
"""

import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

def test_imports():
    """Test that all OneNote modules can be imported."""
    print("Testing OneNote module imports...")

    try:
        from modules.onenote import OneNoteManager, OneNoteCOMClient, OneNoteContentBuilder, TemplateBuilder
        print("✓ OneNote modules imported successfully")
        return True
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False


def test_workflow_helpers():
    """Test that OneNoteHelper is available in workflow_helpers."""
    print("\nTesting workflow helper integration...")

    try:
        from utils.workflow_helpers import OneNoteHelper
        print("✓ OneNoteHelper imported from workflow_helpers")
        return True
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False


def test_content_builder():
    """Test content builder functionality."""
    print("\nTesting content builder...")

    try:
        from modules.onenote import OneNoteContentBuilder, TemplateBuilder

        # Test basic content builder
        builder = OneNoteContentBuilder("Test Page")
        builder.add_heading("Test Heading", 1)
        builder.add_text("Test text content")
        builder.add_bullet_list(["Item 1", "Item 2", "Item 3"])

        content = builder.build_simple()
        assert "Test Page" in content
        assert "Test Heading" in content
        assert "Item 1" in content

        print("✓ Content builder working correctly")

        # Test template builder
        template = TemplateBuilder.meeting_minutes(
            meeting_title="Test Meeting",
            date="2024-12-09",
            attendees=["Person 1", "Person 2"],
            agenda_items=["Topic 1", "Topic 2"],
            notes="Test notes"
        )

        template_content = template.build_simple()
        assert "Test Meeting" in template_content
        assert "Person 1" in template_content

        print("✓ Template builder working correctly")
        return True

    except Exception as e:
        print(f"✗ Content builder test failed: {e}")
        return False


def test_manager_creation():
    """Test OneNote manager creation (without connecting)."""
    print("\nTesting OneNote manager creation...")

    try:
        from modules.onenote import OneNoteManager

        manager = OneNoteManager()
        assert manager.name == "OneNoteManager"
        assert manager.version == "1.0.0"
        assert manager.category == "microsoft_office"

        print("✓ OneNote manager created successfully")
        return True

    except Exception as e:
        print(f"✗ Manager creation failed: {e}")
        return False


def test_com_client_creation():
    """Test COM client creation (without connecting)."""
    print("\nTesting COM client creation...")

    try:
        from modules.onenote import OneNoteCOMClient

        client = OneNoteCOMClient()
        assert client.is_connected() == False

        print("✓ COM client created successfully")
        print("  Note: Actual connection test requires OneNote to be installed")
        return True

    except RuntimeError as e:
        # Expected on non-Windows platforms
        if "pywin32" in str(e):
            print("⚠ COM client requires pywin32 (Windows-only)")
            print("  This is expected on non-Windows platforms")
            print("  On Windows, run: pip install pywin32")
            return True  # Not a failure, just platform-specific
        else:
            print(f"✗ COM client creation failed: {e}")
            return False

    except Exception as e:
        print(f"✗ COM client creation failed: {e}")
        return False


def test_example_file_exists():
    """Test that example file was created."""
    print("\nTesting example file...")

    example_file = Path(__file__).parent / 'examples' / 'onenote_examples.py'

    if example_file.exists():
        print(f"✓ Example file exists: {example_file}")

        # Check file has content
        content = example_file.read_text()
        if len(content) > 1000:
            print(f"✓ Example file has substantial content ({len(content)} chars)")
        return True
    else:
        print(f"✗ Example file not found: {example_file}")
        return False


def test_config_integration():
    """Test that OneNote config can be set/retrieved."""
    print("\nTesting configuration integration...")

    try:
        from core.config import get_config_manager

        config = get_config_manager()

        # Test setting OneNote config
        config.set('onenote.test_notebook', 'Test Notebook')
        config.set('onenote.test_section', 'Test Section')

        # Test retrieving config
        notebook = config.get('onenote.test_notebook')
        section = config.get('onenote.test_section')

        assert notebook == 'Test Notebook'
        assert section == 'Test Section'

        print("✓ Configuration integration working")
        return True

    except Exception as e:
        print(f"✗ Configuration test failed: {e}")
        return False


def main():
    """Run all tests."""
    print("=" * 60)
    print("ONENOTE INTEGRATION TEST SUITE")
    print("=" * 60)
    print()

    results = []

    # Run all tests
    results.append(("Module Imports", test_imports()))
    results.append(("Workflow Helpers", test_workflow_helpers()))
    results.append(("Content Builder", test_content_builder()))
    results.append(("Manager Creation", test_manager_creation()))
    results.append(("COM Client Creation", test_com_client_creation()))
    results.append(("Example File", test_example_file_exists()))
    results.append(("Configuration", test_config_integration()))

    # Print summary
    print()
    print("=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)

    passed = sum(1 for _, result in results if result)
    total = len(results)

    for name, result in results:
        status = "PASS" if result else "FAIL"
        symbol = "✓" if result else "✗"
        print(f"{symbol} {name}: {status}")

    print()
    print(f"Results: {passed}/{total} tests passed")

    if passed == total:
        print("\n🎉 All tests passed! OneNote integration is ready to use.")
        print("\nNext steps:")
        print("1. Open the Automation Hub GUI")
        print("2. Go to Settings > OneNote tab")
        print("3. Configure default notebook and section")
        print("4. Test OneNote connection")
        print("5. Run example workflows from scripts/examples/onenote_examples.py")
        return 0
    else:
        print(f"\n⚠ {total - passed} test(s) failed. Please review the output above.")
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
