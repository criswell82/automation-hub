"""
Schedule Dialog for creating scheduled tasks.
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QComboBox, QDateTimeEdit,
    QSpinBox, QGroupBox, QFormLayout, QCheckBox
)
from PyQt5.QtCore import Qt, QDateTime
from PyQt5.QtGui import QFont

from typing import Dict, Any
from datetime import datetime, timedelta
import uuid

from .script_manager import ScriptMetadata
from .scheduler import ScheduledTask


class ScheduleDialog(QDialog):
    """Dialog for creating scheduled tasks."""

    def __init__(self, script: ScriptMetadata, parameters: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.script = script
        self.parameters = parameters
        self.task = None

        self.setWindowTitle(f"Schedule: {script.name}")
        self.setMinimumWidth(500)

        self._create_ui()

    def _create_ui(self):
        """Create the dialog UI."""
        layout = QVBoxLayout(self)

        # Task name
        name_group = QGroupBox("Task Name")
        name_layout = QFormLayout()
        self.name_edit = QLineEdit()
        self.name_edit.setText(f"{self.script.name} - Scheduled")
        name_layout.addRow("Name:", self.name_edit)
        name_group.setLayout(name_layout)
        layout.addWidget(name_group)

        # Schedule type
        schedule_group = QGroupBox("Schedule Type")
        schedule_layout = QVBoxLayout()

        self.schedule_type = QComboBox()
        self.schedule_type.addItems([
            "One-Time",
            "Daily",
            "Weekly",
            "Every N Hours"
        ])
        self.schedule_type.currentIndexChanged.connect(self._on_schedule_type_changed)
        schedule_layout.addWidget(self.schedule_type)

        schedule_group.setLayout(schedule_layout)
        layout.addWidget(schedule_group)

        # Schedule configuration
        self.config_group = QGroupBox("Schedule Configuration")
        self.config_layout = QFormLayout()
        self.config_group.setLayout(self.config_layout)
        layout.addWidget(self.config_group)

        # Create initial config widgets
        self._create_config_widgets()

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        save_btn = QPushButton("Save Schedule")
        save_btn.clicked.connect(self.accept)
        button_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def _on_schedule_type_changed(self):
        """Handle schedule type change."""
        # Clear existing widgets
        while self.config_layout.count():
            item = self.config_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Create new widgets based on type
        self._create_config_widgets()

    def _create_config_widgets(self):
        """Create configuration widgets based on schedule type."""
        schedule_type = self.schedule_type.currentText()

        if schedule_type == "One-Time":
            self.datetime_edit = QDateTimeEdit()
            self.datetime_edit.setDateTime(QDateTime.currentDateTime().addSecs(3600))
            self.datetime_edit.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
            self.config_layout.addRow("Run Date/Time:", self.datetime_edit)

        elif schedule_type == "Daily":
            self.time_edit = QDateTimeEdit()
            self.time_edit.setDisplayFormat("HH:mm")
            self.time_edit.setTime(QDateTime.currentDateTime().time())
            self.config_layout.addRow("Run Time:", self.time_edit)

        elif schedule_type == "Weekly":
            self.time_edit = QDateTimeEdit()
            self.time_edit.setDisplayFormat("HH:mm")
            self.time_edit.setTime(QDateTime.currentDateTime().time())
            self.config_layout.addRow("Run Time:", self.time_edit)

            self.days_group = QGroupBox("Days of Week")
            days_layout = QHBoxLayout()
            self.day_checkboxes = {}
            for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']:
                cb = QCheckBox(day)
                cb.setChecked(True)
                self.day_checkboxes[day] = cb
                days_layout.addWidget(cb)
            self.days_group.setLayout(days_layout)

            self.config_layout.addRow("", self.days_group)

        elif schedule_type == "Every N Hours":
            self.hours_spin = QSpinBox()
            self.hours_spin.setRange(1, 24)
            self.hours_spin.setValue(1)
            self.config_layout.addRow("Every N Hours:", self.hours_spin)

    def get_scheduled_task(self) -> ScheduledTask:
        """Create scheduled task from dialog inputs."""
        task_id = str(uuid.uuid4())
        name = self.name_edit.text()
        schedule_type_text = self.schedule_type.currentText()

        # Convert UI schedule type to internal format
        schedule_type = 'once'
        schedule_config = {}

        if schedule_type_text == "One-Time":
            schedule_type = 'once'
            dt = self.datetime_edit.dateTime().toPyDateTime()
            schedule_config = {'run_date': dt.isoformat()}

        elif schedule_type_text == "Daily":
            schedule_type = 'cron'
            time = self.time_edit.time()
            schedule_config = {
                'hour': time.hour(),
                'minute': time.minute()
            }

        elif schedule_type_text == "Weekly":
            schedule_type = 'cron'
            time = self.time_edit.time()

            # Get selected days (0=Monday, 6=Sunday)
            selected_days = []
            day_map = {'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3, 'Fri': 4, 'Sat': 5, 'Sun': 6}
            for day, cb in self.day_checkboxes.items():
                if cb.isChecked():
                    selected_days.append(day_map[day])

            schedule_config = {
                'hour': time.hour(),
                'minute': time.minute(),
                'day_of_week': ','.join(map(str, selected_days))
            }

        elif schedule_type_text == "Every N Hours":
            schedule_type = 'interval'
            hours = self.hours_spin.value()
            schedule_config = {'hours': hours}

        task = ScheduledTask(
            id=task_id,
            name=name,
            script_id=self.script.id,
            parameters=self.parameters,
            schedule_type=schedule_type,
            schedule_config=schedule_config,
            enabled=True
        )

        return task
