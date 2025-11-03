"""Validation utility functions."""

import re
from pathlib import Path
from typing import Optional


def validate_email(email: str) -> bool:
    """Validate email address format."""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def validate_file_path(filepath: str, must_exist: bool = False) -> bool:
    """Validate file path."""
    try:
        path = Path(filepath)
        if must_exist:
            return path.exists() and path.is_file()
        return True
    except:
        return False


def validate_url(url: str) -> bool:
    """Validate URL format."""
    pattern = r'^https?://[^\s/$.?#].[^\s]*$'
    return bool(re.match(pattern, url))
