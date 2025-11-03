"""File utility functions."""

import os
from pathlib import Path
from typing import Optional


def ensure_dir(directory: str) -> Path:
    """Ensure a directory exists, create if it doesn't."""
    path = Path(directory)
    path.mkdir(parents=True, exist_ok=True)
    return path


def safe_file_write(filepath: str, content: str, encoding: str = 'utf-8') -> bool:
    """Safely write content to a file."""
    try:
        ensure_dir(os.path.dirname(filepath))
        with open(filepath, 'w', encoding=encoding) as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"Error writing file: {e}")
        return False


def safe_file_read(filepath: str, encoding: str = 'utf-8') -> Optional[str]:
    """Safely read content from a file."""
    try:
        with open(filepath, 'r', encoding=encoding) as f:
            return f.read()
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
