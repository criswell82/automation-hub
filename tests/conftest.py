"""
Shared pytest fixtures for Automation Hub test suite.
"""

import os
import sys
import json
import tempfile
import shutil
from pathlib import Path
from typing import Generator

import pytest

# Add src to path for imports
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for test files."""
    temp_path = Path(tempfile.mkdtemp())
    yield temp_path
    shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def temp_config_dir(temp_dir: Path) -> Path:
    """Create a temporary config directory structure."""
    config_dir = temp_dir / "config"
    config_dir.mkdir(parents=True)
    return config_dir


@pytest.fixture
def sample_config() -> dict:
    """Sample configuration dictionary for testing."""
    return {
        "version": "1.0.0",
        "app": {
            "theme": "dark",
            "language": "en",
            "log_level": "DEBUG"
        },
        "paths": {
            "logs": "/tmp/logs",
            "temp": "/tmp/temp"
        },
        "scheduler": {
            "max_concurrent_jobs": 3,
            "job_timeout_minutes": 30
        }
    }


@pytest.fixture
def sample_template_code() -> str:
    """Sample valid template code for testing."""
    return '''"""
Sample Template

WORKFLOW_META:
name: Sample Template
description: A sample template for testing
category: test
version: 1.0.0
author: Test Author
parameters:
    input_file:
        type: string
        required: true
        description: Path to input file
    output_dir:
        type: string
        required: false
        default: ./output
"""

from modules.base_module import BaseModule


class SampleTemplate(BaseModule):
    def configure(self, **kwargs):
        self.input_file = kwargs.get('input_file')
        self.output_dir = kwargs.get('output_dir', './output')

    def validate(self):
        if not self.input_file:
            raise ValueError("input_file is required")
        return True

    def execute(self):
        return {"status": "success"}


def run(**kwargs):
    template = SampleTemplate("sample")
    template.configure(**kwargs)
    template.validate()
    return template.execute()
'''


@pytest.fixture
def invalid_template_code() -> str:
    """Sample invalid template code (missing metadata)."""
    return '''"""
Invalid Template - No metadata
"""

def run():
    pass
'''


@pytest.fixture
def temp_templates_dir(temp_dir: Path) -> Path:
    """Create a temporary templates directory structure."""
    templates_dir = temp_dir / "templates"
    templates_dir.mkdir(parents=True)
    (templates_dir / "custom").mkdir()
    return templates_dir
