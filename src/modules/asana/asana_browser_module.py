"""
Asana Browser Automation Module

Automates Asana web interface using desktop RPA (browser automation).
Uses existing InputController and WindowManager for UI interaction.

Capabilities:
- Read task lists from projects
- Create new tasks via web forms
- Update task status, assignees, due dates
- Extract project/roadmap data for reporting
- Bulk operations with error handling

Requirements:
- Asana web interface access
- Browser (Chrome/Edge recommended)
- pyautogui for input control
"""

import logging
import time
from typing import List, Dict, Any, Optional
from datetime import datetime
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from modules.base_module import BaseModule, ModuleStatus
from modules.desktop_rpa import InputController, WindowManager

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False


class AsanaBrowserModule(BaseModule):
    """
    Automate Asana via browser RPA.

    Configuration:
        - asana_url: Asana project URL
        - operation: 'read_tasks' | 'create_task' | 'create_tasks' | 'update_task'
        - tasks: Task data (for create/update operations)
        - wait_time: Time to wait for page loads (default 3 seconds)
        - auto_login: Whether to handle login automatically (default False)
    """

    def __init__(self):
        super().__init__(
            name="AsanaBrowserModule",
            description="Automate Asana via browser RPA",
            version="1.0.0",
            category="asana"
        )

        self.input_controller = None
        self.window_manager = None
        self.asana_url = None
        self.operation = None
        self.tasks = []
        self.wait_time = 3
        self.auto_login = False

    def configure(self, **kwargs) -> bool:
        """
        Configure the module.

        Args:
            asana_url: Asana project URL
            operation: Operation to perform
            tasks: Task data (for create/update operations)
            wait_time: Page load wait time in seconds
            auto_login: Handle login automatically
        """
        try:
            # Validate pyautogui availability
            if not PYAUTOGUI_AVAILABLE:
                self.log_error("pyautogui is required but not installed")
                return False

            # Initialize controllers
            self.input_controller = InputController(
                default_delay=kwargs.get('default_delay', 0.5),
                typing_interval=kwargs.get('typing_interval', 0.05)
            )
            self.window_manager = WindowManager()

            # Get configuration
            self.asana_url = kwargs.get('asana_url')
            if not self.asana_url:
                self.log_error("asana_url is required")
                return False

            self.operation = kwargs.get('operation', 'read_tasks')
            self.tasks = kwargs.get('tasks', [])
            self.wait_time = kwargs.get('wait_time', 3)
            self.auto_login = kwargs.get('auto_login', False)

            # Store config
            self.set_config('asana_url', self.asana_url)
            self.set_config('operation', self.operation)
            self.set_config('wait_time', self.wait_time)

            self.log_info(f"Configured for operation: {self.operation}")
            return True

        except Exception as e:
            self.log_error(f"Configuration failed: {e}")
            return False

    def validate(self) -> bool:
        """Validate configuration."""
        try:
            # Validate URL
            if 'asana.com' not in self.asana_url.lower():
                self.log_warning(f"URL doesn't look like an Asana URL: {self.asana_url}")

            # Validate operation
            valid_operations = ['read_tasks', 'create_task', 'create_tasks', 'update_task']
            if self.operation not in valid_operations:
                self.log_error(f"Invalid operation: {self.operation}. Must be one of {valid_operations}")
                return False

            # Validate tasks for create/update operations
            if self.operation in ['create_task', 'create_tasks', 'update_task']:
                if not self.tasks:
                    self.log_error(f"Operation '{self.operation}' requires 'tasks' parameter")
                    return False

            self.log_info("Validation passed")
            return True

        except Exception as e:
            self.log_error(f"Validation failed: {e}")
            return False

    def execute(self) -> Dict[str, Any]:
        """Execute the configured operation."""
        try:
            self.log_info(f"Starting operation: {self.operation}")

            # Open Asana in browser
            self._open_asana()

            # Handle login if needed
            if self.auto_login:
                self._handle_login()

            # Execute operation
            if self.operation == 'read_tasks':
                result = self._read_tasks()
            elif self.operation == 'create_task':
                result = self._create_single_task(self.tasks[0] if self.tasks else {})
            elif self.operation == 'create_tasks':
                result = self._create_multiple_tasks(self.tasks)
            elif self.operation == 'update_task':
                result = self._update_task(self.tasks[0] if self.tasks else {})
            else:
                raise ValueError(f"Unknown operation: {self.operation}")

            return result

        except Exception as e:
            self.log_error(f"Execution failed: {e}")
            return self.error_result(f"Operation failed: {e}")

    def _open_asana(self):
        """Open Asana URL in browser."""
        try:
            self.log_info(f"Opening Asana: {self.asana_url}")

            # Open default browser (you can also use webbrowser module)
            # For now, using keyboard shortcut approach
            import webbrowser
            webbrowser.open(self.asana_url)

            # Wait for page to load
            self.log_info(f"Waiting {self.wait_time} seconds for page to load...")
            time.sleep(self.wait_time)

        except Exception as e:
            self.log_error(f"Failed to open Asana: {e}")
            raise

    def _handle_login(self):
        """
        Handle Asana login if needed.

        Note: This is a placeholder. Actual implementation would need:
        - Credential management (use security module)
        - Login form detection
        - Input of username/password
        - 2FA handling if enabled
        """
        self.log_warning("Auto-login is not fully implemented. Please log in manually.")
        self.log_info("Waiting 10 seconds for manual login...")
        time.sleep(10)

    def _read_tasks(self) -> Dict[str, Any]:
        """
        Read tasks from the current Asana project.

        This is a template implementation. Actual implementation would:
        1. Locate task list elements on page
        2. Extract task names, assignees, due dates
        3. Parse and structure data
        4. Handle pagination

        Returns:
            Result with extracted tasks
        """
        try:
            self.log_info("Reading tasks from Asana...")

            # Wait for page to stabilize
            time.sleep(self.wait_time)

            # Placeholder: In real implementation, would use:
            # - pyautogui.locateOnScreen() to find task list
            # - OCR or screenshot analysis to read text
            # - Systematic scrolling and data collection

            # For now, return a template result
            tasks = []

            self.log_warning("Task reading is a template implementation. Requires customization for your Asana layout.")

            return self.success_result(
                {'tasks': tasks, 'count': len(tasks)},
                f"Read {len(tasks)} tasks (template mode)"
            )

        except Exception as e:
            self.log_error(f"Failed to read tasks: {e}")
            raise

    def _create_single_task(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a single task in Asana via UI automation.

        Steps:
        1. Click "Add task" button (usually Tab or keyboard shortcut)
        2. Type task name
        3. Press Enter to create
        4. Optionally set assignee, due date, etc.

        Args:
            task: Task dictionary with name, description, assignee, due_date

        Returns:
            Result dictionary
        """
        try:
            task_name = task.get('name', 'Untitled Task')
            self.log_info(f"Creating task: {task_name}")

            # Method 1: Use keyboard shortcut (usually Tab or Quick Add)
            # Tab key opens quick add in most Asana layouts
            self.input_controller.press_key('tab', delay=0.5)

            # Wait for input field to appear
            time.sleep(0.5)

            # Type task name
            self.input_controller.type_text(task_name, interval=0.05)

            # Press Enter to create
            self.input_controller.press_key('enter', delay=0.5)

            self.log_info(f"✓ Created task: {task_name}")

            # Optionally set additional fields
            if task.get('description') or task.get('assignee') or task.get('due_date'):
                self.log_info("Setting additional task fields...")
                # Click on task to open details (would need position)
                # Set assignee, due date, description
                # This requires more specific implementation based on Asana layout

            return self.success_result(
                {'task': task_name},
                f"Created task: {task_name}"
            )

        except Exception as e:
            self.log_error(f"Failed to create task: {e}")
            raise

    def _create_multiple_tasks(self, tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Create multiple tasks in bulk.

        Args:
            tasks: List of task dictionaries

        Returns:
            Result with created and failed tasks
        """
        try:
            self.log_info(f"Creating {len(tasks)} tasks...")

            created_tasks = []
            failed_tasks = []

            for i, task in enumerate(tasks, 1):
                self.log_info(f"Creating task {i}/{len(tasks)}: {task.get('name')}")

                try:
                    result = self._create_single_task(task)
                    created_tasks.append(task.get('name'))

                    # Small delay between tasks
                    time.sleep(1)

                except Exception as e:
                    self.log_error(f"Failed to create task: {e}")
                    failed_tasks.append({
                        'name': task.get('name'),
                        'error': str(e)
                    })

            return self.success_result(
                {
                    'created_count': len(created_tasks),
                    'failed_count': len(failed_tasks),
                    'created_tasks': created_tasks,
                    'failed_tasks': failed_tasks
                },
                f"Created {len(created_tasks)} tasks, {len(failed_tasks)} failed"
            )

        except Exception as e:
            self.log_error(f"Bulk task creation failed: {e}")
            raise

    def _update_task(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """
        Update an existing task.

        This would require:
        1. Finding the task by name
        2. Clicking on it to open details
        3. Updating fields (status, assignee, due date, etc.)

        Args:
            task: Task dictionary with updates

        Returns:
            Result dictionary
        """
        try:
            task_name = task.get('name')
            self.log_info(f"Updating task: {task_name}")

            # Placeholder implementation
            self.log_warning("Task update is a template implementation")

            return self.success_result(
                {'task': task_name},
                f"Updated task: {task_name} (template mode)"
            )

        except Exception as e:
            self.log_error(f"Failed to update task: {e}")
            raise

    def take_screenshot(self, filename: str = None) -> str:
        """
        Take a screenshot of current Asana page.

        Useful for:
        - Debugging automation
        - Visual verification
        - OCR text extraction

        Args:
            filename: Optional filename (auto-generated if None)

        Returns:
            Path to screenshot file
        """
        try:
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"asana_screenshot_{timestamp}.png"

            filepath = self.get_output_path(filename)

            screenshot = pyautogui.screenshot()
            screenshot.save(str(filepath))

            self.log_info(f"Screenshot saved: {filepath}")
            return str(filepath)

        except Exception as e:
            self.log_error(f"Failed to take screenshot: {e}")
            raise

    def find_and_click(self, image_path: str, confidence: float = 0.8) -> bool:
        """
        Find an element by image and click it.

        Useful for:
        - Clicking buttons (if you have reference images)
        - Locating UI elements

        Args:
            image_path: Path to reference image
            confidence: Match confidence (0-1)

        Returns:
            True if found and clicked
        """
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)

            if location:
                center = pyautogui.center(location)
                self.input_controller.click(center.x, center.y)
                self.log_info(f"Found and clicked element at {center}")
                return True
            else:
                self.log_warning(f"Element not found on screen: {image_path}")
                return False

        except Exception as e:
            self.log_error(f"Find and click failed: {e}")
            return False


# Example usage
if __name__ == '__main__':
    # Example 1: Create a single task
    module = AsanaBrowserModule()
    result = module.run(
        asana_url='https://app.asana.com/0/1234567890/list',
        operation='create_task',
        tasks=[{
            'name': 'Review project proposal',
            'description': 'Review and provide feedback on Q1 proposal',
            'assignee': 'john.doe'
        }],
        wait_time=3
    )
    print(result)

    # Example 2: Bulk create tasks
    tasks = [
        {'name': 'Task 1: Update documentation'},
        {'name': 'Task 2: Code review'},
        {'name': 'Task 3: Deploy to staging'}
    ]

    module2 = AsanaBrowserModule()
    result2 = module2.run(
        asana_url='https://app.asana.com/0/1234567890/list',
        operation='create_tasks',
        tasks=tasks,
        wait_time=3
    )
    print(result2)
