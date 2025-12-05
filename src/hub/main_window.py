"""
Main Window for Automation Hub.
Central dashboard for managing and executing automation scripts.
"""

import logging
from typing import Optional
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QTextEdit,
    QListWidget, QListWidgetItem, QTreeWidget, QTreeWidgetItem, QSplitter,
    QStatusBar, QMenuBar, QAction, QMessageBox,
    QGroupBox, QScrollArea, QDialog
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

from core.config import get_config_manager
from core.logging_config import get_logging_manager
from .script_manager import ScriptManager, ScriptExecution
from .script_dialog import ScriptDialog
from .scheduler import SchedulerManager, ScheduledTask
from .schedule_dialog import ScheduleDialog
from .workflow_generator_dialog import WorkflowGeneratorDialog
from .settings_dialog import SettingsDialog
from .template_browser_widget import TemplateBrowserWidget
from .template_uploader_widget import TemplateUploaderWidget
from .document_to_template_widget import DocumentToTemplateWidget
from core.template_manager import TemplateManager


class MainWindow(QMainWindow):
    """
    Main application window for Automation Hub.
    """

    def __init__(self) -> None:
        """Initialize the main window."""
        super().__init__()

        self.config_manager = get_config_manager()
        self.logging_manager = get_logging_manager()
        self.logger = logging.getLogger(__name__)

        self.logger.info("Initializing main window...")

        # Initialize script manager
        self.script_manager = ScriptManager()
        self._connect_script_manager()

        # Initialize scheduler
        self.scheduler_manager = SchedulerManager(self.script_manager, self.config_manager)
        self._connect_scheduler()

        # Window setup
        self.setWindowTitle("Automation Hub")
        self.setGeometry(100, 100, 1200, 800)

        # Create UI
        self._create_menu_bar()
        self._create_central_widget()
        self._create_status_bar()

        # Update stats and refresh scheduled tasks
        self._update_dashboard_stats()
        self._refresh_scheduled_tasks()

        self.logger.info("Main window initialized successfully")

    def _connect_script_manager(self) -> None:
        """Connect script manager signals."""
        self.script_manager.execution_started.connect(self._on_execution_started)
        self.script_manager.execution_finished.connect(self._on_execution_finished)
        self.script_manager.output_received.connect(self._on_script_output)

    def _connect_scheduler(self) -> None:
        """Connect scheduler signals."""
        self.scheduler_manager.task_added.connect(self._on_task_added)
        self.scheduler_manager.task_removed.connect(self._on_task_removed)
        self.scheduler_manager.task_executed.connect(self._on_task_executed)

    def _create_menu_bar(self) -> None:
        """Create the menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        new_action = QAction("Create &New Workflow", self)
        new_action.setShortcut("Ctrl+N")
        new_action.setStatusTip("Create a new workflow using AI assistant")
        new_action.triggered.connect(self._on_new_workflow)
        file_menu.addAction(new_action)

        file_menu.addSeparator()

        settings_action = QAction("&Settings", self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self._on_settings)
        file_menu.addAction(settings_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Scripts menu
        scripts_menu = menubar.addMenu("&Scripts")

        run_action = QAction("&Run Script", self)
        run_action.setShortcut("Ctrl+R")
        run_action.triggered.connect(self._on_run_script)
        scripts_menu.addAction(run_action)

        schedule_action = QAction("&Schedule Script", self)
        schedule_action.setShortcut("Ctrl+S")
        schedule_action.triggered.connect(self._on_schedule_script)
        scripts_menu.addAction(schedule_action)

        scripts_menu.addSeparator()

        refresh_action = QAction("Re&fresh Scripts", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(self._on_refresh_scripts)
        scripts_menu.addAction(refresh_action)

        # View menu
        view_menu = menubar.addMenu("&View")

        clear_logs_action = QAction("&Clear Logs", self)
        clear_logs_action.setShortcut("Ctrl+L")
        clear_logs_action.triggered.connect(self._on_clear_logs)
        view_menu.addAction(clear_logs_action)

        # Help menu
        help_menu = menubar.addMenu("&Help")

        docs_action = QAction("&Documentation", self)
        docs_action.setShortcut("F1")
        docs_action.triggered.connect(self._on_show_docs)
        help_menu.addAction(docs_action)

        help_menu.addSeparator()

        about_action = QAction("&About", self)
        about_action.triggered.connect(self._on_about)
        help_menu.addAction(about_action)

    def _create_central_widget(self) -> None:
        """Create the central widget with main content."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        main_layout = QHBoxLayout(central_widget)

        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)

        # Left panel: Scripts list
        left_panel = self._create_scripts_panel()
        splitter.addWidget(left_panel)

        # Right panel: Tabs for different views
        right_panel = self._create_right_panel()
        splitter.addWidget(right_panel)

        # Set initial sizes
        splitter.setSizes([300, 900])

        main_layout.addWidget(splitter)

    def _create_scripts_panel(self) -> QWidget:
        """Create the scripts list panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # Header
        header = QLabel("Automation Scripts")
        header_font = QFont()
        header_font.setPointSize(12)
        header_font.setBold(True)
        header.setFont(header_font)
        layout.addWidget(header)

        # Search/filter (placeholder for future implementation)
        # search_box = QLineEdit()
        # search_box.setPlaceholderText("Search scripts...")
        # layout.addWidget(search_box)

        # Scripts list (tree view with categories)
        self.scripts_list = QTreeWidget()
        self.scripts_list.setHeaderLabel("Script Library")
        self.scripts_list.itemClicked.connect(self._on_script_selected)

        # Populate with example scripts
        self._populate_scripts_list()

        layout.addWidget(self.scripts_list)

        # Action buttons
        button_layout = QHBoxLayout()

        run_btn = QPushButton("Run")
        run_btn.clicked.connect(self._on_run_script)
        button_layout.addWidget(run_btn)

        schedule_btn = QPushButton("Schedule")
        schedule_btn.clicked.connect(self._on_schedule_script)
        button_layout.addWidget(schedule_btn)

        layout.addLayout(button_layout)

        return panel

    def _create_right_panel(self) -> QWidget:
        """Create the right panel with tabs."""
        tabs = QTabWidget()

        # Dashboard tab
        dashboard_tab = self._create_dashboard_tab()
        tabs.addTab(dashboard_tab, "Dashboard")

        # Templates tab
        templates_tab = self._create_templates_tab()
        tabs.addTab(templates_tab, "Templates")

        # Output/Logs tab
        output_tab = self._create_output_tab()
        tabs.addTab(output_tab, "Output")

        # Scheduled Tasks tab
        scheduled_tab = self._create_scheduled_tab()
        tabs.addTab(scheduled_tab, "Scheduled Tasks")

        return tabs

    def _create_dashboard_tab(self) -> QWidget:
        """Create the dashboard tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Welcome message
        welcome = QLabel("Welcome to Automation Hub")
        welcome_font = QFont()
        welcome_font.setPointSize(16)
        welcome_font.setBold(True)
        welcome.setFont(welcome_font)
        welcome.setAlignment(Qt.AlignCenter)
        layout.addWidget(welcome)

        # Quick stats
        stats_group = QGroupBox("Quick Stats")
        stats_layout = QVBoxLayout()

        self.stats_label = QLabel(
            "Total Scripts: 6\n"
            "Last Run: Not yet\n"
            "Scheduled Tasks: 0"
        )
        stats_layout.addWidget(self.stats_label)

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        # Recent activity
        activity_group = QGroupBox("Recent Activity")
        activity_layout = QVBoxLayout()

        self.activity_list = QListWidget()
        self.activity_list.addItem("Application started")
        activity_layout.addWidget(self.activity_list)

        activity_group.setLayout(activity_layout)
        layout.addWidget(activity_group)

        layout.addStretch()

        return widget

    def _create_output_tab(self) -> QWidget:
        """Create the output/logs tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Output display
        self.output_display = QTextEdit()
        self.output_display.setReadOnly(True)
        self.output_display.setFont(QFont("Courier New", 9))
        self.output_display.append("=== Automation Hub Output ===\n")
        self.output_display.append("Ready to execute automation scripts.\n")

        layout.addWidget(self.output_display)

        # Control buttons
        button_layout = QHBoxLayout()

        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._on_clear_logs)
        button_layout.addWidget(clear_btn)

        export_btn = QPushButton("Export Logs")
        export_btn.clicked.connect(self._on_export_logs)
        button_layout.addWidget(export_btn)

        button_layout.addStretch()

        layout.addLayout(button_layout)

        # Set up log monitoring
        self._setup_log_monitoring()

        return widget

    def _create_scheduled_tab(self) -> QWidget:
        """Create the scheduled tasks tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel("Scheduled Tasks")
        label_font = QFont()
        label_font.setPointSize(12)
        label_font.setBold(True)
        label.setFont(label_font)
        layout.addWidget(label)

        self.scheduled_list = QListWidget()
        layout.addWidget(self.scheduled_list)

        # Action buttons
        button_layout = QHBoxLayout()

        add_btn = QPushButton("Add Schedule")
        add_btn.clicked.connect(self._on_schedule_script)
        button_layout.addWidget(add_btn)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh_scheduled_tasks)
        button_layout.addWidget(refresh_btn)

        self.delete_task_btn = QPushButton("Delete")
        self.delete_task_btn.clicked.connect(self._on_delete_task)
        button_layout.addWidget(self.delete_task_btn)

        button_layout.addStretch()

        layout.addLayout(button_layout)

        return widget

    def _create_templates_tab(self) -> QWidget:
        """Create the templates management tab."""
        # Use sub-tabs for Browse, Upload, and Create from Document
        templates_tabs = QTabWidget()

        # Browse Templates sub-tab
        self.template_browser = TemplateBrowserWidget()
        templates_tabs.addTab(self.template_browser, "Browse Templates")

        # Upload Template sub-tab
        self.template_uploader = TemplateUploaderWidget()
        templates_tabs.addTab(self.template_uploader, "Upload Template")

        # Create from Document sub-tab
        self.doc_to_template = DocumentToTemplateWidget()
        templates_tabs.addTab(self.doc_to_template, "Create from Document")

        # Connect signals
        self.template_browser.template_selected.connect(self._on_template_selected)
        self.template_browser.template_exported.connect(self._on_template_exported)
        self.template_uploader.template_imported.connect(self._on_template_imported)
        self.doc_to_template.template_created.connect(self._on_template_created)

        return templates_tabs

    def _create_status_bar(self) -> None:
        """Create the status bar."""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def _populate_scripts_list(self) -> None:
        """Populate the scripts list with available automation modules, grouped by category."""
        scripts = self.script_manager.get_scripts()

        # Group scripts by category
        scripts_by_category = {}
        for script in scripts:
            category = script.category
            if category not in scripts_by_category:
                scripts_by_category[category] = []
            scripts_by_category[category].append(script)

        # Create tree items for each category
        for category in sorted(scripts_by_category.keys()):
            # Create category header
            category_item = QTreeWidgetItem()
            category_item.setText(0, f"{category} ({len(scripts_by_category[category])})")
            category_item.setData(0, Qt.UserRole, None)  # Mark as category header (not executable)

            # Make category header bold
            font = category_item.font(0)
            font.setBold(True)
            category_item.setFont(0, font)

            # Add scripts under this category
            for script in scripts_by_category[category]:
                script_item = QTreeWidgetItem(category_item)
                script_item.setText(0, script.name)
                script_item.setData(0, Qt.UserRole, script.id)  # Store script ID
                script_item.setToolTip(0, script.description)  # Show description on hover

            self.scripts_list.addTopLevelItem(category_item)

        # Expand all categories by default
        self.scripts_list.expandAll()

    def _setup_log_monitoring(self) -> None:
        """Setup periodic log monitoring."""
        self.log_timer = QTimer()
        self.log_timer.timeout.connect(self._update_logs)
        self.log_timer.start(2000)  # Update every 2 seconds

    def _update_logs(self) -> None:
        """Update the output display with recent logs."""
        # Get recent logs from logging manager
        recent_logs = self.logging_manager.get_recent_logs(lines=10)

        if recent_logs and len(recent_logs) > 0:
            # Only show new logs
            current_text = self.output_display.toPlainText()
            new_logs = ''.join(recent_logs)

            if new_logs not in current_text:
                self.output_display.append(new_logs)
                # Auto-scroll to bottom
                self.output_display.verticalScrollBar().setValue(
                    self.output_display.verticalScrollBar().maximum()
                )

    # Slot methods
    def _on_script_selected(self, item: QTreeWidgetItem):
        """Handle script selection."""
        # Check if this is a category header (script_id will be None)
        script_id = item.data(0, Qt.UserRole)

        if script_id is None:
            # Category header was clicked, just return (allow expand/collapse)
            return

        script_name = item.text(0)
        self.logger.info(f"Selected script: {script_name}")
        self.status_bar.showMessage(f"Selected: {script_name}")

    def _on_run_script(self):
        """Handle run script action."""
        current_item = self.scripts_list.currentItem()

        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select a script to run.")
            return

        # Get script from manager
        script_id = current_item.data(0, Qt.UserRole)

        # Check if a category was selected instead of a script
        if script_id is None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a specific script, not a category.")
            return

        script = self.script_manager.get_script(script_id)

        if not script:
            QMessageBox.warning(self, "Error", "Script not found!")
            return

        # Open script dialog
        dialog = ScriptDialog(script, self)

        if dialog.exec_() == dialog.Accepted:
            # Get parameters
            parameters = dialog.get_parameters()

            if dialog.is_dry_run():
                # Dry run - just show parameters
                params_str = "\n".join([f"  {k}: {v}" for k, v in parameters.items()])
                QMessageBox.information(
                    self,
                    "Dry Run",
                    f"Script: {script.name}\n\nParameters:\n{params_str}\n\n"
                    "This was a dry run. Click 'Run' without 'Dry Run' checked to execute."
                )
            else:
                # Actually execute
                self.script_manager.execute_script(script, parameters)

    def _on_schedule_script(self):
        """Handle schedule script action."""
        current_item = self.scripts_list.currentItem()

        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select a script to schedule.")
            return

        # Get script from manager
        script_id = current_item.data(0, Qt.UserRole)

        # Check if a category was selected instead of a script
        if script_id is None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a specific script, not a category.")
            return

        script = self.script_manager.get_script(script_id)

        if not script:
            QMessageBox.warning(self, "Error", "Script not found!")
            return

        # First, get parameters using the script dialog
        script_dialog = ScriptDialog(script, self)

        if script_dialog.exec_() == script_dialog.Accepted:
            parameters = script_dialog.get_parameters()

            # Now open schedule dialog
            schedule_dialog = ScheduleDialog(script, parameters, self)

            if schedule_dialog.exec_() == schedule_dialog.Accepted:
                task = schedule_dialog.get_scheduled_task()

                if self.scheduler_manager.add_task(task):
                    QMessageBox.information(
                        self,
                        "Task Scheduled",
                        f"Task '{task.name}' has been scheduled successfully.\n\n"
                        f"Schedule: {task.schedule_type}\n"
                        f"Next run: {task.next_run.strftime('%Y-%m-%d %H:%M:%S') if task.next_run else 'Pending'}"
                    )
                    # Refresh scheduled tasks list
                    self._refresh_scheduled_tasks()
                else:
                    QMessageBox.critical(
                        self,
                        "Scheduling Failed",
                        "Failed to schedule the task. Check logs for details."
                    )

    def _on_refresh_scripts(self):
        """Handle refresh scripts action."""
        self.logger.info("Refreshing scripts list...")
        self.scripts_list.clear()
        self._populate_scripts_list()

        # Also refresh template browser if it exists
        if hasattr(self, 'template_browser'):
            self.template_browser.refresh()

        self.status_bar.showMessage("Scripts and templates refreshed")

    def _on_new_workflow(self):
        """Handle create new workflow action."""
        try:
            dialog = WorkflowGeneratorDialog(self)
            dialog.workflow_created.connect(self._on_workflow_created)
            dialog.exec_()
        except Exception as e:
            self.logger.error(f"Failed to open workflow generator: {e}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to open workflow generator:\n\n{str(e)}"
            )

    def _on_workflow_created(self, file_path: str):
        """Handle workflow created event."""
        self.logger.info(f"New workflow created: {file_path}")

        # Refresh scripts to include the new workflow
        self._on_refresh_scripts()

        # Show success message
        self.status_bar.showMessage(f"Workflow created: {Path(file_path).name}", 5000)

        # Optionally switch to scripts tab to show the new workflow
        if hasattr(self, 'tabs'):
            self.tabs.setCurrentIndex(0)  # Switch to dashboard/scripts view

    def _on_template_selected(self, template):
        """Handle template selection from browser."""
        try:
            # Open the workflow generator dialog with the selected template
            dialog = WorkflowGeneratorDialog(self)
            dialog.workflow_created.connect(self._on_workflow_created)

            # Load the template into the generator
            dialog._on_template_selected(template)

            dialog.exec_()
        except Exception as e:
            self.logger.error(f"Failed to load template: {e}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to load template:\n\n{str(e)}"
            )

    def _on_template_exported(self, file_path: str):
        """Handle template export event."""
        self.logger.info(f"Template exported to: {file_path}")
        self.status_bar.showMessage(f"Template exported: {Path(file_path).name}", 5000)

    def _on_template_imported(self, file_path: str):
        """Handle template import event."""
        self.logger.info(f"Template imported: {file_path}")

        # Refresh scripts to include the new template
        self._on_refresh_scripts()

        # Refresh template browser to show the new template
        if hasattr(self, 'template_browser'):
            self.template_browser.refresh()

        # Show success message
        self.status_bar.showMessage(f"Template imported: {Path(file_path).name}", 5000)

    def _on_template_created(self, file_path: str):
        """Handle template created from document event."""
        self.logger.info(f"Template created from document: {file_path}")

        # Refresh scripts to include the new template
        self._on_refresh_scripts()

        # Refresh template browser to show the new template
        if hasattr(self, 'template_browser'):
            self.template_browser.refresh()

        # Show success message
        self.status_bar.showMessage(f"Template created: {Path(file_path).name}", 5000)

    def _on_settings(self):
        """Handle settings action."""
        try:
            dialog = SettingsDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                self.status_bar.showMessage("Settings saved successfully", 3000)
                self.logger.info("Settings updated")
        except Exception as e:
            self.logger.error(f"Failed to open settings dialog: {e}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to open settings:\n\n{str(e)}"
            )

    def _on_clear_logs(self):
        """Handle clear logs action."""
        self.output_display.clear()
        self.output_display.append("=== Logs Cleared ===\n")
        self.status_bar.showMessage("Logs cleared")

    def _on_export_logs(self):
        """Handle export logs action."""
        # Export logs to file
        log_file = self.logging_manager.archive_logs()
        if log_file:
            QMessageBox.information(
                self,
                "Logs Exported",
                f"Logs exported to:\n{log_file}"
            )
        else:
            QMessageBox.warning(
                self,
                "Export Failed",
                "Failed to export logs."
            )

    def _on_show_docs(self):
        """Handle show documentation action."""
        QMessageBox.information(
            self,
            "Documentation",
            "Documentation will be displayed here.\n\n"
            "In the full version, this will open the user guide."
        )

    def _on_about(self):
        """Handle about action."""
        QMessageBox.about(
            self,
            "About Automation Hub",
            "<h2>Automation Hub</h2>"
            "<p>Version 1.0.0-alpha</p>"
            "<p>Corporate Desktop Automation Platform</p>"
            "<p>&copy; 2025 Automation Hub Development Team</p>"
            "<p>Built with PyQt5 and Python</p>"
        )

    def _on_execution_started(self, execution: ScriptExecution):
        """Handle script execution start."""
        self.output_display.append(f"\n{'=' * 60}")
        self.output_display.append(f">>> Executing: {execution.script.name}")
        self.output_display.append(f"Started at: {execution.start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        self.output_display.append(f"{'=' * 60}\n")

        self.activity_list.insertItem(0, f"▶ Running: {execution.script.name}")
        self.status_bar.showMessage(f"Executing: {execution.script.name}")

    def _on_execution_finished(self, execution: ScriptExecution):
        """Handle script execution completion."""
        duration = execution.get_duration()
        status_symbol = "✓" if execution.status == 'success' else "✗"

        self.output_display.append(f"\n{'-' * 60}")
        self.output_display.append(f"{status_symbol} Execution {execution.status.upper()}")
        self.output_display.append(f"Duration: {duration:.2f} seconds")
        if execution.error:
            self.output_display.append(f"Error: {execution.error}")
        self.output_display.append(f"{'-' * 60}\n")

        self.activity_list.insertItem(0, f"{status_symbol} {execution.script.name}: {execution.status}")
        self.status_bar.showMessage(f"Execution {execution.status}: {execution.script.name}")

        # Update dashboard stats
        self._update_dashboard_stats()

        # Show result dialog for errors
        if execution.error:
            QMessageBox.critical(
                self,
                "Execution Error",
                f"Script: {execution.script.name}\n\n"
                f"Error: {execution.error}\n\n"
                f"Duration: {duration:.2f} seconds"
            )
        else:
            QMessageBox.information(
                self,
                "Execution Complete",
                f"Script: {execution.script.name}\n\n"
                f"Status: Success\n"
                f"Duration: {duration:.2f} seconds"
            )

    def _on_script_output(self, output: str):
        """Handle output from script execution."""
        self.output_display.append(output)
        # Auto-scroll to bottom
        self.output_display.verticalScrollBar().setValue(
            self.output_display.verticalScrollBar().maximum()
        )

    def _update_dashboard_stats(self) -> None:
        """Update dashboard statistics."""
        stats = self.script_manager.get_statistics()
        history = self.script_manager.get_execution_history(limit=1)

        last_run = "Never"
        if history:
            last_run = history[-1].start_time.strftime('%Y-%m-%d %H:%M:%S')

        tasks_count = len(self.scheduler_manager.get_tasks())

        self.stats_label.setText(
            f"Total Scripts: {len(self.script_manager.get_scripts())}\n"
            f"Total Executions: {stats['total_executions']}\n"
            f"Successful: {stats['successful']}\n"
            f"Failed: {stats['failed']}\n"
            f"Success Rate: {stats['success_rate']:.1f}%\n"
            f"Scheduled Tasks: {tasks_count}\n"
            f"Last Run: {last_run}"
        )

    def _refresh_scheduled_tasks(self) -> None:
        """Refresh the scheduled tasks list."""
        self.scheduled_list.clear()

        tasks = self.scheduler_manager.get_tasks()

        if not tasks:
            self.scheduled_list.addItem("No scheduled tasks")
            return

        for task in tasks:
            status = "✓" if task.enabled else "✗"
            next_run = task.next_run.strftime('%Y-%m-%d %H:%M') if task.next_run else 'Pending'

            item_text = f"{status} {task.name} | Next: {next_run} | Runs: {task.run_count}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, task.id)
            self.scheduled_list.addItem(item)

    def _on_delete_task(self):
        """Handle delete scheduled task."""
        current_item = self.scheduled_list.currentItem()

        if not current_item:
            return

        task_id = current_item.data(Qt.UserRole)
        task = self.scheduler_manager.get_task(task_id)

        if not task:
            return

        reply = QMessageBox.question(
            self,
            "Delete Task",
            f"Are you sure you want to delete the scheduled task:\n\n'{task.name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            if self.scheduler_manager.remove_task(task_id):
                self._refresh_scheduled_tasks()
                self._update_dashboard_stats()
                QMessageBox.information(self, "Success", "Task deleted successfully")

    def _on_task_added(self, task: ScheduledTask):
        """Handle task added event."""
        self._refresh_scheduled_tasks()
        self._update_dashboard_stats()
        self.activity_list.insertItem(0, f"⏰ Scheduled: {task.name}")

    def _on_task_removed(self, task_id: str):
        """Handle task removed event."""
        self._refresh_scheduled_tasks()
        self._update_dashboard_stats()

    def _on_task_executed(self, task_id: str, success: bool):
        """Handle scheduled task execution."""
        task = self.scheduler_manager.get_task(task_id)
        if task:
            status = "✓" if success else "✗"
            self.activity_list.insertItem(0, f"⏰{status} {task.name}: {'Success' if success else 'Failed'}")
            self._refresh_scheduled_tasks()

    def closeEvent(self, event):
        """Handle window close event."""
        self.logger.info("Closing Automation Hub...")

        reply = QMessageBox.question(
            self,
            "Confirm Exit",
            "Are you sure you want to exit Automation Hub?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Shutdown scheduler
            self.scheduler_manager.shutdown()
            self.logger.info("Application closed by user")
            event.accept()
        else:
            event.ignore()
