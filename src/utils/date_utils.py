"""Date and time utility functions."""

from datetime import datetime
from typing import Optional


def format_datetime(dt: datetime, format_str: str = "%Y-%m-%d %H:%M:%S") -> str:
    """Format a datetime object."""
    return dt.strftime(format_str)


def parse_datetime(date_str: str, format_str: str = "%Y-%m-%d %H:%M:%S") -> Optional[datetime]:
    """Parse a datetime string."""
    try:
        return datetime.strptime(date_str, format_str)
    except ValueError:
        return None


def get_timestamp(format_str: str = "%Y%m%d_%H%M%S") -> str:
    """Get current timestamp as string."""
    return datetime.now().strftime(format_str)
