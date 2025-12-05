"""
Workflow Helpers

Common automation patterns and helper functions for building workflows.

This module re-exports helpers from specialized modules for backwards compatibility.
For new code, prefer importing directly from the specialized modules:
- utils.report_helpers
- utils.onenote_helpers
- utils.communication_helpers
- utils.file_helpers
- utils.asana_helpers
"""

# Re-export all public interfaces for backwards compatibility

# Report helpers
from utils.report_helpers import (
    ReportBuilder,
    aggregate_data_from_files,
    create_pivot_summary,
)

# OneNote helpers
from utils.onenote_helpers import OneNoteHelper

# Communication helpers
from utils.communication_helpers import (
    EmailHelper,
    run_daily_report_workflow,
)

# File helpers
from utils.file_helpers import (
    FileOrganizer,
    SharePointHelper,
    run_file_processing_workflow,
)

# Asana helpers
from utils.asana_helpers import (
    AsanaHelper,
    create_asana_roadmap_from_excel,
    sync_asana_status_to_excel,
)

__all__ = [
    # Report
    'ReportBuilder',
    'aggregate_data_from_files',
    'create_pivot_summary',
    # OneNote
    'OneNoteHelper',
    # Communication
    'EmailHelper',
    'run_daily_report_workflow',
    # File
    'FileOrganizer',
    'SharePointHelper',
    'run_file_processing_workflow',
    # Asana
    'AsanaHelper',
    'create_asana_roadmap_from_excel',
    'sync_asana_status_to_excel',
]
