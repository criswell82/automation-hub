"""
Unit tests for TemplateManager.
"""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from core.template_manager import TemplateManager, Template, ValidationResult


class TestTemplateManager:
    """Tests for TemplateManager class."""

    @pytest.fixture
    def template_manager(self, temp_templates_dir: Path) -> TemplateManager:
        """Create a TemplateManager with temp directory."""
        with patch('core.template_manager.get_logger') as mock_logger:
            mock_logger.return_value = MagicMock()
            return TemplateManager(templates_dir=str(temp_templates_dir))

    def test_init_creates_custom_dir(self, temp_templates_dir: Path):
        """Test that TemplateManager creates custom directory."""
        with patch('core.template_manager.get_logger') as mock_logger:
            mock_logger.return_value = MagicMock()
            manager = TemplateManager(templates_dir=str(temp_templates_dir))

        assert (temp_templates_dir / "custom").exists()

    def test_discover_templates_empty(self, template_manager: TemplateManager):
        """Test discovering templates in empty directory."""
        templates = template_manager.discover_templates()

        assert templates == []

    def test_discover_templates_with_valid_template(
        self,
        template_manager: TemplateManager,
        temp_templates_dir: Path,
        sample_template_code: str
    ):
        """Test discovering a valid template."""
        # Create a template file
        template_dir = temp_templates_dir / "test_category"
        template_dir.mkdir()
        template_file = template_dir / "sample_template.py"
        template_file.write_text(sample_template_code)

        templates = template_manager.discover_templates()

        assert len(templates) == 1
        assert templates[0].name == "Sample Template"
        assert templates[0].category == "test"

    def test_validate_template_valid(
        self,
        template_manager: TemplateManager,
        sample_template_code: str
    ):
        """Test validation of valid template code."""
        result = template_manager.validate_template(sample_template_code)

        assert result.valid is True
        assert len(result.errors) == 0
        assert result.metadata is not None
        assert result.metadata['name'] == 'Sample Template'

    def test_validate_template_missing_metadata(
        self,
        template_manager: TemplateManager,
        invalid_template_code: str
    ):
        """Test validation fails for template without WORKFLOW_META."""
        result = template_manager.validate_template(invalid_template_code)

        assert result.valid is False
        assert "Missing WORKFLOW_META" in result.errors[0]

    def test_validate_template_missing_run_function(
        self,
        template_manager: TemplateManager
    ):
        """Test validation fails for template without run() function."""
        code = '''"""
Template

WORKFLOW_META:
name: No Run Function
description: Missing run function
category: test
"""

class SomeClass:
    pass
'''
        result = template_manager.validate_template(code)

        assert result.valid is False
        assert any("run()" in error for error in result.errors)

    def test_validate_template_missing_required_fields(
        self,
        template_manager: TemplateManager
    ):
        """Test validation fails for template missing required fields."""
        code = '''"""
Template

WORKFLOW_META:
name: Only Name
"""

def run():
    pass
'''
        result = template_manager.validate_template(code)

        assert result.valid is False
        assert any("description" in error for error in result.errors)
        assert any("category" in error for error in result.errors)

    def test_save_template(
        self,
        template_manager: TemplateManager,
        sample_template_code: str
    ):
        """Test saving a valid template."""
        result = template_manager.save_template(
            code=sample_template_code,
            category="custom",
            filename="saved_template.py"
        )

        assert result is not None
        assert Path(result).exists()
        assert "saved_template.py" in result

    def test_save_template_auto_filename(
        self,
        template_manager: TemplateManager,
        sample_template_code: str
    ):
        """Test saving template with auto-generated filename."""
        result = template_manager.save_template(
            code=sample_template_code,
            category="custom"
        )

        assert result is not None
        assert "sample_template.py" in result

    def test_save_template_invalid_fails(
        self,
        template_manager: TemplateManager,
        invalid_template_code: str
    ):
        """Test saving invalid template fails."""
        result = template_manager.save_template(
            code=invalid_template_code,
            category="custom"
        )

        assert result is None

    def test_save_template_avoids_overwrite(
        self,
        template_manager: TemplateManager,
        sample_template_code: str
    ):
        """Test saving template with same name creates numbered version."""
        # Save first template
        result1 = template_manager.save_template(
            code=sample_template_code,
            category="custom",
            filename="duplicate.py"
        )

        # Save second template with same name
        result2 = template_manager.save_template(
            code=sample_template_code,
            category="custom",
            filename="duplicate.py"
        )

        assert result1 != result2
        assert "duplicate_1.py" in result2

    def test_load_template(
        self,
        template_manager: TemplateManager,
        temp_templates_dir: Path,
        sample_template_code: str
    ):
        """Test loading a template from file."""
        # Create template file
        template_file = temp_templates_dir / "custom" / "load_test.py"
        template_file.write_text(sample_template_code)

        template = template_manager.load_template(str(template_file))

        assert template is not None
        assert template.name == "Sample Template"
        assert template.code == sample_template_code

    def test_load_template_nonexistent(self, template_manager: TemplateManager):
        """Test loading nonexistent template returns None."""
        template = template_manager.load_template("/nonexistent/template.py")

        assert template is None

    def test_import_template(
        self,
        template_manager: TemplateManager,
        temp_dir: Path,
        sample_template_code: str
    ):
        """Test importing template from external file."""
        # Create external template
        external_file = temp_dir / "external_template.py"
        external_file.write_text(sample_template_code)

        result = template_manager.import_template(
            source_path=str(external_file),
            category="custom"
        )

        assert result is not None
        assert Path(result).exists()

    def test_import_template_invalid(
        self,
        template_manager: TemplateManager,
        temp_dir: Path,
        invalid_template_code: str
    ):
        """Test importing invalid template fails."""
        external_file = temp_dir / "invalid_external.py"
        external_file.write_text(invalid_template_code)

        result = template_manager.import_template(str(external_file))

        assert result is None

    def test_get_template_categories(
        self,
        template_manager: TemplateManager,
        temp_templates_dir: Path,
        sample_template_code: str
    ):
        """Test getting list of template categories."""
        # Create templates in different categories
        for category in ["email", "files", "reports"]:
            cat_dir = temp_templates_dir / category
            cat_dir.mkdir(exist_ok=True)

            code = sample_template_code.replace(
                "category: test",
                f"category: {category}"
            )
            (cat_dir / f"{category}_template.py").write_text(code)

        categories = template_manager.get_template_categories()

        assert "email" in categories
        assert "files" in categories
        assert "reports" in categories


class TestTemplate:
    """Tests for Template dataclass."""

    def test_template_creation(self):
        """Test creating a Template instance."""
        template = Template(
            id="test_template",
            name="Test Template",
            description="A test template",
            category="test",
            file_path="/path/to/template.py",
            parameters={"param1": {"type": "string"}}
        )

        assert template.id == "test_template"
        assert template.name == "Test Template"
        assert template.version == "1.0.0"  # default
        assert template.is_system is True  # default

    def test_template_with_code(self):
        """Test Template with code field."""
        template = Template(
            id="test",
            name="Test",
            description="Test",
            category="test",
            file_path="/path.py",
            parameters={},
            code="def run(): pass"
        )

        assert template.code == "def run(): pass"


class TestValidationResult:
    """Tests for ValidationResult dataclass."""

    def test_valid_result(self):
        """Test creating a valid ValidationResult."""
        result = ValidationResult(
            valid=True,
            errors=[],
            warnings=["Minor warning"],
            metadata={"name": "Test"}
        )

        assert result.valid is True
        assert len(result.errors) == 0
        assert len(result.warnings) == 1

    def test_invalid_result(self):
        """Test creating an invalid ValidationResult."""
        result = ValidationResult(
            valid=False,
            errors=["Missing required field"],
            warnings=[]
        )

        assert result.valid is False
        assert "Missing required field" in result.errors
