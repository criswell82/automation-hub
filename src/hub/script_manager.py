"""
Script Manager for Automation Hub.
Handles script discovery, loading, execution, and tracking.
"""

import logging
import importlib
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional, Callable
from datetime import datetime
import threading
import queue
import io
import traceback

from PyQt5.QtCore import QObject, pyqtSignal, QThread


class ScriptMetadata:
    """Metadata for an automation script."""

    def __init__(
            self,
            id: str,
            name: str,
            description: str,
            category: str,
            module_path: str,
            parameters: Dict[str, Any] = None,
            example_path: Optional[str] = None
    ):
        self.id = id
        self.name = name
        self.description = description
        self.category = category
        self.module_path = module_path
        self.parameters = parameters or {}
        self.example_path = example_path

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'id': self.id,
            'name': self.name,
            'description': self.description,
            'category': self.category,
            'module_path': self.module_path,
            'parameters': self.parameters,
            'example_path': self.example_path
        }


class ScriptExecution:
    """Represents a script execution."""

    def __init__(self, script: ScriptMetadata, parameters: Dict[str, Any]):
        self.script = script
        self.parameters = parameters
        self.start_time = datetime.now()
        self.end_time: Optional[datetime] = None
        self.status = 'running'
        self.output: List[str] = []
        self.error: Optional[str] = None
        self.result: Optional[Dict[str, Any]] = None

    def complete(self, result: Dict[str, Any], error: Optional[str] = None):
        """Mark execution as complete."""
        self.end_time = datetime.now()
        self.result = result
        self.error = error
        self.status = 'error' if error else 'success'

    def get_duration(self) -> float:
        """Get execution duration in seconds."""
        if self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return (datetime.now() - self.start_time).total_seconds()


class ScriptExecutor(QThread):
    """Thread for executing scripts."""

    output_signal = pyqtSignal(str)  # Emits output lines
    finished_signal = pyqtSignal(dict, str)  # Emits (result, error)
    progress_signal = pyqtSignal(int)  # Emits progress percentage

    def __init__(self, script: ScriptMetadata, parameters: Dict[str, Any]):
        super().__init__()
        self.script = script
        self.parameters = parameters
        self.logger = logging.getLogger(__name__)

    def run(self):
        """Execute the script."""
        result = None
        error = None

        try:
            # Capture stdout
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            stdout_capture = io.StringIO()
            stderr_capture = io.StringIO()

            sys.stdout = stdout_capture
            sys.stderr = stderr_capture

            try:
                # Import and execute based on category
                result = self._execute_script()

            finally:
                # Restore stdout/stderr
                sys.stdout = old_stdout
                sys.stderr = old_stderr

                # Emit captured output
                stdout_text = stdout_capture.getvalue()
                stderr_text = stderr_capture.getvalue()

                if stdout_text:
                    self.output_signal.emit(stdout_text)
                if stderr_text:
                    self.output_signal.emit(f"STDERR: {stderr_text}")

        except Exception as e:
            error = f"Execution failed: {str(e)}\n{traceback.format_exc()}"
            self.logger.error(error)

        self.finished_signal.emit(result or {}, error or "")

    def _execute_script(self) -> Dict[str, Any]:
        """Execute the specific script based on category."""
        category = self.script.category.lower()

        if category == 'desktop_rpa':
            return self._execute_rpa_script()
        elif category == 'excel':
            return self._execute_excel_script()
        elif category == 'outlook':
            return self._execute_outlook_script()
        elif category == 'sharepoint':
            return self._execute_sharepoint_script()
        elif category == 'word':
            return self._execute_word_script()
        elif category == 'onenote':
            return self._execute_onenote_script()
        else:
            return {'status': 'error', 'message': f'Unknown category: {category}'}

    def _execute_rpa_script(self) -> Dict[str, Any]:
        """Execute Desktop RPA script."""
        print(f"Executing Desktop RPA script: {self.script.name}")
        print(f"Parameters: {self.parameters}")

        # For now, just run the example if available
        if self.script.example_path:
            print(f"\nExample script available at: {self.script.example_path}")
            print("To run this script, execute it directly with Python.")

        return {
            'status': 'success',
            'message': 'Desktop RPA module ready. Use the example scripts to automate tasks.',
            'category': 'desktop_rpa'
        }

    def _execute_excel_script(self) -> Dict[str, Any]:
        """Execute Excel automation script."""
        print(f"Executing Excel script: {self.script.name}")
        print(f"Parameters: {self.parameters}")

        from modules.excel_automation import WorkbookHandler

        # Example: Create a simple workbook
        if 'output_file' in self.parameters:
            wb = WorkbookHandler()
            sheet = wb.workbook.active
            sheet.title = "Automation Test"

            # Write sample data
            data = [
                ['Name', 'Value', 'Status'],
                ['Test 1', 100, 'Complete'],
                ['Test 2', 200, 'In Progress'],
                ['Test 3', 300, 'Complete']
            ]
            wb.write_data('Automation Test', data, 'A1')

            # Format header
            wb.format_cells('Automation Test', 'A1:C1', bold=True, fill_color='4472C4')

            wb.save(self.parameters['output_file'])
            print(f"\nWorkbook saved to: {self.parameters['output_file']}")

            return {'status': 'success', 'message': 'Excel file created successfully'}

        return {'status': 'success', 'message': 'Excel module ready'}

    def _execute_outlook_script(self) -> Dict[str, Any]:
        """Execute Outlook automation script."""
        print(f"Executing Outlook script: {self.script.name}")
        return {'status': 'success', 'message': 'Outlook module ready'}

    def _execute_sharepoint_script(self) -> Dict[str, Any]:
        """Execute SharePoint script."""
        print(f"Executing SharePoint script: {self.script.name}")
        return {'status': 'success', 'message': 'SharePoint module ready'}

    def _execute_word_script(self) -> Dict[str, Any]:
        """Execute Word automation script."""
        print(f"Executing Word script: {self.script.name}")
        return {'status': 'success', 'message': 'Word module ready'}

    def _execute_onenote_script(self) -> Dict[str, Any]:
        """Execute OneNote script."""
        print(f"Executing OneNote script: {self.script.name}")
        return {'status': 'success', 'message': 'OneNote module ready'}


class ScriptManager(QObject):
    """
    Manages automation scripts - discovery, execution, and tracking.
    """

    execution_started = pyqtSignal(ScriptExecution)
    execution_finished = pyqtSignal(ScriptExecution)
    output_received = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.logger = logging.getLogger(__name__)
        self.scripts: List[ScriptMetadata] = []
        self.executions: List[ScriptExecution] = []
        self.current_executor: Optional[ScriptExecutor] = None

        # Discover scripts
        self.discover_scripts()

    def discover_scripts(self):
        """Discover available automation scripts."""
        self.logger.info("Discovering automation scripts...")

        # Define available scripts
        self.scripts = [
            ScriptMetadata(
                id='desktop_rpa_window',
                name='Desktop RPA - Window Automation',
                description='Automate Windows applications using window management and input control',
                category='Desktop_RPA',
                module_path='modules.desktop_rpa',
                parameters={
                    'window_title': {'type': 'string', 'required': False, 'description': 'Target window title'},
                    'action': {'type': 'choice', 'choices': ['activate', 'move', 'resize'], 'default': 'activate'}
                },
                example_path='src/modules/desktop_rpa/examples/example_notepad_automation.py'
            ),
            ScriptMetadata(
                id='desktop_rpa_copy_paste',
                name='Desktop RPA - Copy & Paste',
                description='Copy data from one application and paste into another',
                category='Desktop_RPA',
                module_path='modules.desktop_rpa',
                parameters={},
                example_path='src/modules/desktop_rpa/examples/example_copy_paste_automation.py'
            ),
            ScriptMetadata(
                id='excel_data_processing',
                name='Excel - Data Processing',
                description='Process and manipulate Excel data with formulas and formatting',
                category='Excel',
                module_path='modules.excel_automation',
                parameters={
                    'input_file': {'type': 'file', 'required': False, 'description': 'Input Excel file'},
                    'output_file': {'type': 'string', 'required': False, 'description': 'Output Excel file path'}
                }
            ),
            ScriptMetadata(
                id='excel_report_generation',
                name='Excel - Report Generation',
                description='Generate formatted reports with charts and pivot tables',
                category='Excel',
                module_path='modules.excel_automation',
                parameters={
                    'template': {'type': 'file', 'required': False, 'description': 'Report template'},
                    'output_file': {'type': 'string', 'required': False, 'description': 'Output file path'}
                }
            ),
            ScriptMetadata(
                id='outlook_email',
                name='Outlook - Email Automation',
                description='Send emails, process inbox, and manage tasks',
                category='Outlook',
                module_path='modules.outlook_automation',
                parameters={
                    'to': {'type': 'string', 'required': False, 'description': 'Recipient email'},
                    'subject': {'type': 'string', 'required': False, 'description': 'Email subject'},
                    'body': {'type': 'text', 'required': False, 'description': 'Email body'}
                }
            ),
            ScriptMetadata(
                id='sharepoint_file_management',
                name='SharePoint - File Management',
                description='Upload, download, and manage SharePoint files',
                category='SharePoint',
                module_path='modules.sharepoint',
                parameters={
                    'site_url': {'type': 'string', 'required': False, 'description': 'SharePoint site URL'},
                    'action': {'type': 'choice', 'choices': ['upload', 'download', 'list'], 'default': 'list'}
                }
            ),
            ScriptMetadata(
                id='word_document',
                name='Word - Document Generation',
                description='Create and format Word documents from templates',
                category='Word',
                module_path='modules.word_automation',
                parameters={
                    'template': {'type': 'file', 'required': False, 'description': 'Document template'},
                    'output_file': {'type': 'string', 'required': False, 'description': 'Output file path'}
                }
            )
        ]

        self.logger.info(f"Discovered {len(self.scripts)} scripts")

    def get_scripts(self, category: Optional[str] = None) -> List[ScriptMetadata]:
        """Get all scripts, optionally filtered by category."""
        if category:
            return [s for s in self.scripts if s.category.lower() == category.lower()]
        return self.scripts

    def get_script(self, script_id: str) -> Optional[ScriptMetadata]:
        """Get a script by ID."""
        for script in self.scripts:
            if script.id == script_id:
                return script
        return None

    def execute_script(self, script: ScriptMetadata, parameters: Dict[str, Any]):
        """Execute a script with parameters."""
        self.logger.info(f"Executing script: {script.name}")

        # Create execution record
        execution = ScriptExecution(script, parameters)
        self.executions.append(execution)

        # Create executor thread
        self.current_executor = ScriptExecutor(script, parameters)

        # Connect signals
        self.current_executor.output_signal.connect(self._on_output)
        self.current_executor.finished_signal.connect(
            lambda result, error: self._on_finished(execution, result, error)
        )

        # Emit started signal
        self.execution_started.emit(execution)

        # Start execution
        self.current_executor.start()

    def _on_output(self, output: str):
        """Handle output from script execution."""
        self.output_received.emit(output)

    def _on_finished(self, execution: ScriptExecution, result: Dict[str, Any], error: str):
        """Handle script execution completion."""
        execution.complete(result, error)
        self.execution_finished.emit(execution)

        self.logger.info(f"Script execution finished: {execution.status}")

    def get_execution_history(self, limit: int = 10) -> List[ScriptExecution]:
        """Get recent execution history."""
        return self.executions[-limit:] if self.executions else []

    def get_statistics(self) -> Dict[str, Any]:
        """Get execution statistics."""
        total = len(self.executions)
        success = len([e for e in self.executions if e.status == 'success'])
        errors = len([e for e in self.executions if e.status == 'error'])

        return {
            'total_executions': total,
            'successful': success,
            'failed': errors,
            'success_rate': (success / total * 100) if total > 0 else 0
        }
