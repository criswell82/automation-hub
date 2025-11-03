"""
Main Window for Automation Hub.
Central dashboard for managing and executing automation scripts.
"""

import logging
from typing import Optional
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QTextEdit,
    QListWidget, QListWidgetItem, QSplitter,
    QStatusBar, QMenuBar, QAction, QMessageBox,
    QGroupBox, QScrollArea
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

from core.config import get_config_manager
from core.logging_config import get_logging_manager
from .script_manager import ScriptManager, ScriptExecution
from .script_dialog import ScriptDialog


class MainWindow(QMainWindow):
    """
    Main application window for Automation Hub.
    """

    def __init__(self):
        """Initialize the main window."""
        super().__init__()

        self.config_manager = get_config_manager()
        self.logging_manager = get_logging_manager()
        self.logger = logging.getLogger(__name__)

        self.logger.info("Initializing main window...")

        # Initialize script manager
        self.script_manager = ScriptManager()
        self._connect_script_manager()

        # Window setup
        self.setWindowTitle("Automation Hub")
        self.setGeometry(100, 100, 1200, 800)

        # Create UI
        self._create_menu_bar()
        self._create_central_widget()
        self._create_status_bar()

        # Update stats
        self._update_dashboard_stats()

        self.logger.info("Main window initialized successfully")

    def _connect_script_manager(self):
        """Connect script manager signals."""
        self.script_manager.execution_started.connect(self._on_execution_started)
        self.script_manager.execution_finished.connect(self._on_execution_finished)
        self.script_manager.output_received.connect(self._on_script_output)

    def _create_menu_bar(self):
        """Create the menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        new_action = QAction("&New Script", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self._on_new_script)
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

    def _create_central_widget(self):
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

        # Scripts list
        self.scripts_list = QListWidget()
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
        self.scheduled_list.addItem("No scheduled tasks")
        layout.addWidget(self.scheduled_list)

        # Action buttons
        button_layout = QHBoxLayout()

        add_btn = QPushButton("Add Schedule")
        add_btn.clicked.connect(self._on_schedule_script)
        button_layout.addWidget(add_btn)

        edit_btn = QPushButton("Edit")
        edit_btn.setEnabled(False)
        button_layout.addWidget(edit_btn)

        delete_btn = QPushButton("Delete")
        delete_btn.setEnabled(False)
        button_layout.addWidget(delete_btn)

        button_layout.addStretch()

        layout.addLayout(button_layout)

        return widget

    def _create_status_bar(self):
        """Create the status bar."""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def _populate_scripts_list(self):
        """Populate the scripts list with available automation modules."""
        scripts = self.script_manager.get_scripts()

        for script in scripts:
            item = QListWidgetItem(script.name)
            item.setData(Qt.UserRole, script.id)  # Store script ID
            self.scripts_list.addItem(item)

    def _setup_log_monitoring(self):
        """Setup periodic log monitoring."""
        self.log_timer = QTimer()
        self.log_timer.timeout.connect(self._update_logs)
        self.log_timer.start(2000)  # Update every 2 seconds

    def _update_logs(self):
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
    def _on_script_selected(self, item: QListWidgetItem):
        """Handle script selection."""
        script_name = item.text()
        self.logger.info(f"Selected script: {script_name}")
        self.status_bar.showMessage(f"Selected: {script_name}")

    def _on_run_script(self):
        """Handle run script action."""
        current_item = self.scripts_list.currentItem()

        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select a script to run.")
            return

        # Get script from manager
        script_id = current_item.data(Qt.UserRole)
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
        QMessageBox.information(
            self,
            "Schedule Script",
            "Task scheduling dialog will be implemented here.\n\n"
            "Will integrate with APScheduler for recurring tasks."
        )

    def _on_refresh_scripts(self):
        """Handle refresh scripts action."""
        self.logger.info("Refreshing scripts list...")
        self.scripts_list.clear()
        self._populate_scripts_list()
        self.status_bar.showMessage("Scripts refreshed")

    def _on_new_script(self):
        """Handle new script action."""
        QMessageBox.information(
            self,
            "New Script",
            "New script wizard will be implemented here."
        )

    def _on_settings(self):
        """Handle settings action."""
        QMessageBox.information(
            self,
            "Settings",
            "Settings dialog will be implemented here.\n\n"
            "Will include:\n"
            "- Application preferences\n"
            "- Module configuration\n"
            "- Logging settings\n"
            "- Credential management"
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

    def _update_dashboard_stats(self):
        """Update dashboard statistics."""
        stats = self.script_manager.get_statistics()
        history = self.script_manager.get_execution_history(limit=1)

        last_run = "Never"
        if history:
            last_run = history[-1].start_time.strftime('%Y-%m-%d %H:%M:%S')

        self.stats_label.setText(
            f"Total Scripts: {len(self.script_manager.get_scripts())}\n"
            f"Total Executions: {stats['total_executions']}\n"
            f"Successful: {stats['successful']}\n"
            f"Failed: {stats['failed']}\n"
            f"Success Rate: {stats['success_rate']:.1f}%\n"
            f"Last Run: {last_run}"
        )

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
            self.logger.info("Application closed by user")
            event.accept()
        else:
            event.ignore()
