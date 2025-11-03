"""
Script Execution Dialog for Automation Hub.
Provides UI for configuring and running scripts.
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QTextEdit, QComboBox,
    QFileDialog, QGroupBox, QFormLayout, QCheckBox,
    QProgressBar, QDialogButtonBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from typing import Dict, Any
from .script_manager import ScriptMetadata


class ScriptDialog(QDialog):
    """Dialog for configuring and executing scripts."""

    def __init__(self, script: ScriptMetadata, parent=None):
        super().__init__(parent)
        self.script = script
        self.parameters = {}
        self.parameter_widgets = {}

        self.setWindowTitle(f"Run Script: {script.name}")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)

        self._create_ui()

    def _create_ui(self):
        """Create the dialog UI."""
        layout = QVBoxLayout(self)

        # Script info
        info_group = self._create_info_section()
        layout.addWidget(info_group)

        # Parameters
        if self.script.parameters:
            param_group = self._create_parameters_section()
            layout.addWidget(param_group)
        else:
            no_params_label = QLabel("This script has no configurable parameters.")
            no_params_label.setStyleSheet("color: gray; font-style: italic;")
            layout.addWidget(no_params_label)

        # Progress bar (hidden initially)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Buttons
        button_layout = self._create_buttons()
        layout.addLayout(button_layout)

    def _create_info_section(self) -> QGroupBox:
        """Create script information section."""
        group = QGroupBox("Script Information")
        layout = QVBoxLayout()

        # Name
        name_label = QLabel(self.script.name)
        name_font = QFont()
        name_font.setPointSize(11)
        name_font.setBold(True)
        name_label.setFont(name_font)
        layout.addWidget(name_label)

        # Description
        desc_label = QLabel(self.script.description)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #555;")
        layout.addWidget(desc_label)

        # Category
        category_label = QLabel(f"Category: {self.script.category}")
        category_label.setStyleSheet("color: #007ACC; font-weight: bold;")
        layout.addWidget(category_label)

        # Example path if available
        if self.script.example_path:
            example_label = QLabel(f"Example: {self.script.example_path}")
            example_label.setStyleSheet("font-size: 9pt; color: gray;")
            example_label.setWordWrap(True)
            layout.addWidget(example_label)

        group.setLayout(layout)
        return group

    def _create_parameters_section(self) -> QGroupBox:
        """Create parameters input section."""
        group = QGroupBox("Parameters")
        form_layout = QFormLayout()

        for param_name, param_info in self.script.parameters.items():
            widget = self._create_parameter_widget(param_name, param_info)
            self.parameter_widgets[param_name] = widget

            # Create label
            label_text = param_name.replace('_', ' ').title()
            if param_info.get('required', False):
                label_text += " *"

            label = QLabel(label_text)

            # Add description as tooltip
            if 'description' in param_info:
                widget.setToolTip(param_info['description'])
                label.setToolTip(param_info['description'])

            form_layout.addRow(label, widget)

        group.setLayout(form_layout)
        return group

    def _create_parameter_widget(self, param_name: str, param_info: Dict[str, Any]):
        """Create appropriate widget for parameter type."""
        param_type = param_info.get('type', 'string')

        if param_type == 'string':
            widget = QLineEdit()
            if 'default' in param_info:
                widget.setText(str(param_info['default']))
            return widget

        elif param_type == 'text':
            widget = QTextEdit()
            widget.setMaximumHeight(100)
            if 'default' in param_info:
                widget.setPlainText(str(param_info['default']))
            return widget

        elif param_type == 'file':
            container = QHBoxLayout()
            file_edit = QLineEdit()
            browse_btn = QPushButton("Browse...")
            browse_btn.clicked.connect(
                lambda: self._browse_file(file_edit)
            )
            container.addWidget(file_edit)
            container.addWidget(browse_btn)

            widget_container = QWidget()
            widget_container.setLayout(container)
            widget_container.file_edit = file_edit  # Store reference
            return widget_container

        elif param_type == 'choice':
            widget = QComboBox()
            choices = param_info.get('choices', [])
            widget.addItems(choices)
            if 'default' in param_info:
                index = widget.findText(param_info['default'])
                if index >= 0:
                    widget.setCurrentIndex(index)
            return widget

        elif param_type == 'boolean':
            widget = QCheckBox()
            widget.setChecked(param_info.get('default', False))
            return widget

        else:
            # Default to string input
            return QLineEdit()

    def _browse_file(self, line_edit: QLineEdit):
        """Open file browser dialog."""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select File",
            "",
            "All Files (*.*)"
        )
        if filename:
            line_edit.setText(filename)

    def _create_buttons(self) -> QHBoxLayout:
        """Create dialog buttons."""
        layout = QHBoxLayout()

        # Dry run checkbox
        self.dry_run_checkbox = QCheckBox("Dry Run (test without execution)")
        layout.addWidget(self.dry_run_checkbox)

        layout.addStretch()

        # Run button
        self.run_btn = QPushButton("Run")
        self.run_btn.setDefault(True)
        self.run_btn.clicked.connect(self.accept)
        layout.addWidget(self.run_btn)

        # Cancel button
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(cancel_btn)

        return layout

    def get_parameters(self) -> Dict[str, Any]:
        """Get parameter values from widgets."""
        parameters = {}

        for param_name, widget in self.parameter_widgets.items():
            if isinstance(widget, QLineEdit):
                parameters[param_name] = widget.text()
            elif isinstance(widget, QTextEdit):
                parameters[param_name] = widget.toPlainText()
            elif isinstance(widget, QComboBox):
                parameters[param_name] = widget.currentText()
            elif isinstance(widget, QCheckBox):
                parameters[param_name] = widget.isChecked()
            elif hasattr(widget, 'file_edit'):
                # File browser widget
                parameters[param_name] = widget.file_edit.text()

        return parameters

    def is_dry_run(self) -> bool:
        """Check if dry run mode is enabled."""
        return self.dry_run_checkbox.isChecked()

    def show_progress(self, value: int):
        """Show progress bar with value."""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(value)

    def hide_progress(self):
        """Hide progress bar."""
        self.progress_bar.setVisible(False)


from PyQt5.QtWidgets import QWidget
