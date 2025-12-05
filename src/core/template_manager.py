"""
Template Manager

Handles template discovery, validation, loading, and saving for workflow templates.
"""

import re
import yaml
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass

from core.logging_config import get_logger

logger = get_logger(__name__)


@dataclass
class Template:
    """Represents a workflow template"""
    id: str
    name: str
    description: str
    category: str
    file_path: str
    parameters: Dict[str, Any]
    version: str = "1.0.0"
    author: str = "Unknown"
    code: Optional[str] = None
    is_system: bool = True  # System template vs user template


@dataclass
class ValidationResult:
    """Result of template validation"""
    valid: bool
    errors: List[str]
    warnings: List[str]
    metadata: Optional[Dict[str, Any]] = None


class TemplateManager:
    """Manages workflow templates"""

    def __init__(self, templates_dir: str = None):
        """
        Initialize template manager

        Args:
            templates_dir: Path to templates directory (defaults to user_scripts/templates/)
        """
        if templates_dir is None:
            project_root = Path(__file__).parent.parent.parent
            templates_dir = project_root / "user_scripts" / "templates"

        self.templates_dir = Path(templates_dir)
        self.custom_dir = self.templates_dir / "custom"

        # Ensure custom directory exists
        self.custom_dir.mkdir(parents=True, exist_ok=True)

        logger.info(f"Template manager initialized: {self.templates_dir}")

    def discover_templates(self) -> List[Template]:
        """
        Discover all available templates

        Returns:
            List of Template objects
        """
        templates = []

        if not self.templates_dir.exists():
            logger.warning(f"Templates directory not found: {self.templates_dir}")
            return templates

        # Scan template directories
        for template_dir in self.templates_dir.iterdir():
            if not template_dir.is_dir():
                continue

            # Recursively find .py files
            for py_file in template_dir.rglob("*.py"):
                if py_file.name.startswith("_"):
                    continue  # Skip private files

                try:
                    template = self._load_template_metadata(py_file)
                    if template:
                        # Mark as system or user template
                        template.is_system = not str(py_file).startswith(str(self.custom_dir))
                        templates.append(template)
                        logger.debug(f"Loaded template: {template.name}")
                except Exception as e:
                    logger.warning(f"Failed to load template {py_file.name}: {e}")

        logger.info(f"Discovered {len(templates)} templates")
        return templates

    def load_template(self, file_path: str) -> Optional[Template]:
        """
        Load a template from file path

        Args:
            file_path: Path to template file

        Returns:
            Template object or None if loading fails
        """
        try:
            path = Path(file_path)
            if not path.exists():
                logger.error(f"Template file not found: {file_path}")
                return None

            template = self._load_template_metadata(path)

            # Also load the full code
            if template:
                with open(path, 'r', encoding='utf-8') as f:
                    template.code = f.read()

            return template
        except Exception as e:
            logger.error(f"Failed to load template from {file_path}: {e}")
            return None

    def validate_template(self, code: str) -> ValidationResult:
        """
        Validate template code

        Args:
            code: Template Python code

        Returns:
            ValidationResult with validation status and messages
        """
        errors = []
        warnings = []
        metadata = None

        # Check for WORKFLOW_META
        docstring_pattern = r'("""|\'\'\')\s*\n?(.*?)WORKFLOW_META:?\s*\n(.*?)("""|\'\'\')'
        match = re.search(docstring_pattern, code, re.DOTALL)

        if not match:
            errors.append("Missing WORKFLOW_META in docstring")
            return ValidationResult(valid=False, errors=errors, warnings=warnings)

        # Parse WORKFLOW_META
        try:
            meta_content = match.group(3).rstrip()
            metadata = yaml.safe_load(meta_content)
        except Exception as e:
            errors.append(f"Invalid WORKFLOW_META YAML: {e}")
            return ValidationResult(valid=False, errors=errors, warnings=warnings)

        # Validate required fields
        required_fields = ['name', 'description', 'category']
        for field in required_fields:
            if field not in metadata:
                errors.append(f"Missing required field: {field}")

        # Check for BaseModule import
        if 'BaseModule' not in code:
            warnings.append("Template should import and extend BaseModule")

        # Check for run() function
        if 'def run(' not in code:
            errors.append("Template must have a run() function")

        # Check for configure(), validate(), execute() methods
        if 'def configure(' not in code:
            warnings.append("Template should have configure() method")
        if 'def validate(' not in code:
            warnings.append("Template should have validate() method")
        if 'def execute(' not in code:
            warnings.append("Template should have execute() method")

        # Check parameters format
        if 'parameters' in metadata:
            params = metadata['parameters']
            if not isinstance(params, dict):
                errors.append("parameters must be a dictionary")
            else:
                # Validate parameter definitions
                for param_name, param_def in params.items():
                    if isinstance(param_def, dict):
                        if 'type' not in param_def:
                            warnings.append(f"Parameter '{param_name}' missing type field")

        valid = len(errors) == 0
        return ValidationResult(valid=valid, errors=errors, warnings=warnings, metadata=metadata)

    def save_template(self, code: str, category: str, filename: str = None) -> Optional[str]:
        """
        Save a template to the templates directory

        Args:
            code: Template Python code
            category: Template category (custom, reports, email, files)
            filename: Optional filename (will be auto-generated from name if not provided)

        Returns:
            Path to saved file or None if save failed
        """
        try:
            # Validate first
            validation = self.validate_template(code)
            if not validation.valid:
                logger.error(f"Cannot save invalid template: {validation.errors}")
                return None

            # Determine target directory
            if category.lower() == 'custom':
                target_dir = self.custom_dir
            else:
                target_dir = self.templates_dir / category.lower()

            target_dir.mkdir(parents=True, exist_ok=True)

            # Generate filename if not provided
            if not filename:
                # Extract name from metadata
                name = validation.metadata.get('name', 'template')
                filename = name.lower().replace(' ', '_').replace('-', '_') + '.py'

            # Ensure .py extension
            if not filename.endswith('.py'):
                filename += '.py'

            target_path = target_dir / filename

            # Check if file already exists
            if target_path.exists():
                # Add number suffix to avoid overwriting
                base_name = target_path.stem
                counter = 1
                while target_path.exists():
                    target_path = target_dir / f"{base_name}_{counter}.py"
                    counter += 1

            # Save file
            with open(target_path, 'w', encoding='utf-8') as f:
                f.write(code)

            logger.info(f"Template saved to: {target_path}")
            return str(target_path)

        except Exception as e:
            logger.error(f"Failed to save template: {e}")
            return None

    def import_template(self, source_path: str, category: str = 'custom') -> Optional[str]:
        """
        Import a template from an external file

        Args:
            source_path: Path to source template file
            category: Category to import into (default: custom)

        Returns:
            Path to imported template or None if import failed
        """
        try:
            source = Path(source_path)
            if not source.exists():
                logger.error(f"Source file not found: {source_path}")
                return None

            # Read and validate template
            with open(source, 'r', encoding='utf-8') as f:
                code = f.read()

            validation = self.validate_template(code)
            if not validation.valid:
                logger.error(f"Invalid template: {validation.errors}")
                return None

            # Save to templates directory
            return self.save_template(code, category, source.name)

        except Exception as e:
            logger.error(f"Failed to import template: {e}")
            return None

    def export_template(self, template_id: str, dest_path: str) -> bool:
        """
        Export a template to an external file

        Args:
            template_id: ID of template to export
            dest_path: Destination path

        Returns:
            True if export succeeded
        """
        try:
            # Find template
            templates = self.discover_templates()
            template = next((t for t in templates if t.id == template_id), None)

            if not template:
                logger.error(f"Template not found: {template_id}")
                return False

            # Copy file
            dest = Path(dest_path)
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(template.file_path, dest)

            logger.info(f"Template exported to: {dest_path}")
            return True

        except Exception as e:
            logger.error(f"Failed to export template: {e}")
            return False

    def delete_template(self, template_id: str) -> bool:
        """
        Delete a template (only custom templates can be deleted)

        Args:
            template_id: ID of template to delete

        Returns:
            True if deletion succeeded
        """
        try:
            # Find template
            templates = self.discover_templates()
            template = next((t for t in templates if t.id == template_id), None)

            if not template:
                logger.error(f"Template not found: {template_id}")
                return False

            # Only allow deleting custom templates
            if template.is_system:
                logger.error("Cannot delete system templates")
                return False

            # Delete file
            Path(template.file_path).unlink()
            logger.info(f"Template deleted: {template.name}")
            return True

        except Exception as e:
            logger.error(f"Failed to delete template: {e}")
            return False

    def _load_template_metadata(self, file_path: Path) -> Optional[Template]:
        """
        Load template metadata from file

        Args:
            file_path: Path to template file

        Returns:
            Template object or None if loading fails
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Parse WORKFLOW_META
        docstring_pattern = r'("""|\'\'\')\s*\n?(.*?)WORKFLOW_META:?\s*\n(.*?)("""|\'\'\')'
        match = re.search(docstring_pattern, content, re.DOTALL)

        if not match:
            return None

        try:
            meta_content = match.group(3).rstrip()
            metadata = yaml.safe_load(meta_content)
        except Exception as e:
            logger.warning(f"Failed to parse metadata from {file_path.name}: {e}")
            return None

        # Generate template ID
        try:
            rel_path = file_path.relative_to(self.templates_dir)
            template_id = str(rel_path.with_suffix('')).replace('\\', '_').replace('/', '_')
        except ValueError:
            template_id = file_path.stem

        return Template(
            id=f"template_{template_id}",
            name=metadata.get('name', file_path.stem.replace('_', ' ').title()),
            description=metadata.get('description', 'No description provided'),
            category=metadata.get('category', 'Custom'),
            file_path=str(file_path),
            parameters=metadata.get('parameters', {}),
            version=metadata.get('version', '1.0.0'),
            author=metadata.get('author', 'Unknown')
        )

    def get_template_categories(self) -> List[str]:
        """Get list of available template categories"""
        categories = set()
        templates = self.discover_templates()
        for template in templates:
            categories.add(template.category)
        return sorted(list(categories))
