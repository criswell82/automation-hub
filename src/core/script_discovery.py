"""
Script Discovery System

Scans user_scripts directory for custom workflows and dynamically loads them.
"""

import os
import re
import ast
import importlib.util
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
import yaml

from core.logging_config import get_logger

logger = get_logger(__name__)


@dataclass
class ScriptMetadata:
    """Metadata for a discovered script"""
    id: str
    name: str
    description: str
    category: str
    module_path: str
    file_path: str
    parameters: Dict[str, Any]
    version: str = "1.0.0"
    author: str = "Unknown"
    tags: List[str] = None

    def __post_init__(self):
        if self.tags is None:
            self.tags = []


class ScriptDiscovery:
    """Discovers and loads custom workflow scripts"""

    def __init__(self, scripts_dir: str = None):
        """
        Initialize script discovery

        Args:
            scripts_dir: Path to user_scripts directory (defaults to user_scripts/)
        """
        if scripts_dir is None:
            # Get project root and construct path to user_scripts
            project_root = Path(__file__).parent.parent.parent
            scripts_dir = project_root / "user_scripts"

        self.scripts_dir = Path(scripts_dir)
        self.discovered_scripts: List[ScriptMetadata] = []
        logger.info(f"Script discovery initialized: {self.scripts_dir}")

    def scan_directory(self) -> List[ScriptMetadata]:
        """
        Scan the user_scripts directory for workflow scripts

        Returns:
            List of ScriptMetadata objects
        """
        self.discovered_scripts.clear()

        if not self.scripts_dir.exists():
            logger.warning(f"Scripts directory not found: {self.scripts_dir}")
            return []

        # Scan custom, template, and test directories
        search_paths = [
            self.scripts_dir / "custom",
            self.scripts_dir / "templates" / "reports",
            self.scripts_dir / "templates" / "email",
            self.scripts_dir / "templates" / "files",
            self.scripts_dir / "templates" / "custom",  # User-uploaded templates
            self.scripts_dir / "tests",
        ]

        for search_path in search_paths:
            if search_path.exists():
                self._scan_path(search_path)

        logger.info(f"Discovered {len(self.discovered_scripts)} custom scripts")
        return self.discovered_scripts

    def _scan_path(self, path: Path):
        """Scan a specific path for Python scripts"""
        for py_file in path.rglob("*.py"):
            if py_file.name.startswith("_"):
                continue  # Skip private/internal files

            try:
                metadata = self._extract_metadata(py_file)
                if metadata:
                    self.discovered_scripts.append(metadata)
                    logger.debug(f"Loaded script: {metadata.name}")
            except Exception as e:
                logger.warning(f"Failed to load {py_file.name}: {e}")

    def _extract_metadata(self, file_path: Path) -> Optional[ScriptMetadata]:
        """
        Extract metadata from a Python script file

        Args:
            file_path: Path to the Python file

        Returns:
            ScriptMetadata object or None if metadata not found
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Look for WORKFLOW_META in docstring
        metadata = self._parse_docstring_metadata(content)
        if not metadata:
            return None

        # Generate unique ID from file path
        script_id = self._generate_script_id(file_path)

        # Construct module path for import
        module_path = self._construct_module_path(file_path)

        return ScriptMetadata(
            id=script_id,
            name=metadata.get('name', file_path.stem.replace('_', ' ').title()),
            description=metadata.get('description', 'No description provided'),
            category=metadata.get('category', 'Custom'),
            module_path=module_path,
            file_path=str(file_path),
            parameters=metadata.get('parameters', {}),
            version=metadata.get('version', '1.0.0'),
            author=metadata.get('author', 'Unknown'),
            tags=metadata.get('tags', [])
        )

    def _parse_docstring_metadata(self, content: str) -> Optional[Dict]:
        """
        Parse WORKFLOW_META from docstring

        Supports formats:
        1. YAML in docstring:
           '''
           WORKFLOW_META:
             name: My Workflow
             description: Does something
           '''

        2. Python dict in docstring:
           '''
           WORKFLOW_META = {
             'name': 'My Workflow',
             'description': 'Does something'
           }
           '''
        """
        # Try to find docstring with WORKFLOW_META (supports both """ and ''')
        docstring_pattern = r'("""|\'\'\')\s*\n?(.*?)WORKFLOW_META:?\s*\n(.*?)("""|\'\'\')'
        match = re.search(docstring_pattern, content, re.DOTALL)

        if not match:
            return None

        # Group 3 contains the YAML content after WORKFLOW_META:
        # Only strip trailing whitespace, preserve leading indentation for YAML parsing
        meta_content = match.group(3).rstrip()

        # Try to parse as YAML
        try:
            # The content should already be valid YAML, try parsing directly first
            metadata = yaml.safe_load(meta_content)
            if metadata:
                return metadata

            # If direct parsing fails, try extracting YAML block
            yaml_content = self._extract_yaml_block(meta_content)
            metadata = yaml.safe_load(yaml_content)
            return metadata
        except Exception as e:
            logger.debug(f"YAML parsing failed: {e}")

        # Try to parse as Python dict
        try:
            # Look for dictionary pattern
            dict_pattern = r'\{.*?\}'
            dict_match = re.search(dict_pattern, meta_content, re.DOTALL)
            if dict_match:
                metadata = ast.literal_eval(dict_match.group(0))
                return metadata
        except Exception as e:
            logger.debug(f"Dict parsing failed: {e}")

        return None

    def _extract_yaml_block(self, text: str) -> str:
        """Extract indented YAML block"""
        lines = text.split('\n')
        yaml_lines = []

        for line in lines:
            # Stop at unindented line or closing quotes
            if line and not line[0].isspace() and line.strip() != '':
                break
            yaml_lines.append(line)

        return '\n'.join(yaml_lines)

    def _generate_script_id(self, file_path: Path) -> str:
        """Generate unique script ID from file path"""
        # Use relative path from user_scripts as ID
        try:
            rel_path = file_path.relative_to(self.scripts_dir)
            script_id = str(rel_path.with_suffix('')).replace(os.sep, '_')
            return f"custom_{script_id}"
        except ValueError:
            # Fallback to filename
            return f"custom_{file_path.stem}"

    def _construct_module_path(self, file_path: Path) -> str:
        """Construct Python module path for import"""
        # Return the file path for dynamic import
        return str(file_path.absolute())

    def load_script_module(self, metadata: ScriptMetadata):
        """
        Dynamically load a script module

        Args:
            metadata: ScriptMetadata object

        Returns:
            Loaded module object
        """
        try:
            spec = importlib.util.spec_from_file_location(
                metadata.id,
                metadata.module_path
            )
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                return module
            else:
                raise ImportError(f"Could not load spec for {metadata.id}")
        except Exception as e:
            logger.error(f"Failed to load module {metadata.id}: {e}")
            raise

    def execute_script(self, metadata: ScriptMetadata, **kwargs) -> Any:
        """
        Execute a discovered script

        Args:
            metadata: ScriptMetadata object
            **kwargs: Parameters to pass to the script

        Returns:
            Script execution result
        """
        try:
            module = self.load_script_module(metadata)

            # Look for run() function
            if hasattr(module, 'run'):
                return module.run(**kwargs)

            # Fallback: look for class that inherits from BaseModule
            elif hasattr(module, metadata.name.replace(' ', '')):
                workflow_class = getattr(module, metadata.name.replace(' ', ''))
                workflow = workflow_class()
                workflow.configure(**kwargs)
                workflow.validate()
                return workflow.execute()

            else:
                raise AttributeError(
                    f"Script must have run() function or {metadata.name.replace(' ', '')} class"
                )

        except Exception as e:
            logger.error(f"Failed to execute script {metadata.name}: {e}")
            raise

    def get_script_by_id(self, script_id: str) -> Optional[ScriptMetadata]:
        """Get script metadata by ID"""
        for script in self.discovered_scripts:
            if script.id == script_id:
                return script
        return None

    def get_scripts_by_category(self, category: str) -> List[ScriptMetadata]:
        """Get all scripts in a category"""
        return [s for s in self.discovered_scripts if s.category == category]

    def refresh(self) -> List[ScriptMetadata]:
        """Refresh the script list (rescan directory)"""
        logger.info("Refreshing script library...")
        return self.scan_directory()


# Singleton instance
_script_discovery = None


def get_script_discovery(scripts_dir: str = None) -> ScriptDiscovery:
    """Get the global ScriptDiscovery instance"""
    global _script_discovery
    if _script_discovery is None:
        _script_discovery = ScriptDiscovery(scripts_dir)
    return _script_discovery


if __name__ == "__main__":
    # Test script discovery
    discovery = ScriptDiscovery()
    scripts = discovery.scan_directory()

    print(f"\nDiscovered {len(scripts)} scripts:")
    for script in scripts:
        print(f"  - {script.name} ({script.category})")
        print(f"    ID: {script.id}")
        print(f"    File: {script.file_path}")
        print(f"    Parameters: {list(script.parameters.keys())}")
        print()
