"""
Template Uploader Widget

UI for uploading and importing workflow templates.
"""

import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTextEdit, QComboBox, QFileDialog,
    QMessageBox, QGroupBox, QFormLayout
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent

from core.template_manager import TemplateManager, ValidationResult


class TemplateUploaderWidget(QWidget):
    """Widget for uploading and importing templates"""

    template_imported = pyqtSignal(str)  # Emits path to imported template

    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.template_manager = TemplateManager()
        self.selected_file = None
        self.validation_result = None

        self._create_ui()

    def _create_ui(self):
        """Create the uploader UI"""
        layout = QVBoxLayout(self)

        # Instructions
        instructions = QLabel(
            "Upload workflow template files (.py) to add them to your template library.\n"
            "Templates will be validated before import."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        # File selection group
        file_group = QGroupBox("Select Template File")
        file_layout = QVBoxLayout()

        # Drag-drop area
        self.drop_area = QLabel("Drag and drop template file here\n\nOR")
        self.drop_area.setAlignment(Qt.AlignCenter)
        self.drop_area.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 5px;
                padding: 40px;
                background-color: #f5f5f5;
                min-height: 100px;
            }
        """)
        self.drop_area.setAcceptDrops(True)
        self.drop_area.dragEnterEvent = self._drag_enter
        self.drop_area.dropEvent = self._drop
        file_layout.addWidget(self.drop_area)

        # Browse button
        browse_btn = QPushButton("Browse for Template File...")
        browse_btn.clicked.connect(self._browse_file)
        file_layout.addWidget(browse_btn)

        # Selected file display
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: #666; font-style: italic;")
        file_layout.addWidget(self.file_label)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Options group
        options_group = QGroupBox("Import Options")
        options_layout = QFormLayout()

        # Category selection
        self.category_combo = QComboBox()
        self.category_combo.addItems(['Custom', 'Reports', 'Email', 'Files'])
        self.category_combo.setCurrentText('Custom')
        options_layout.addRow("Category:", self.category_combo)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Validation results group
        validation_group = QGroupBox("Validation Results")
        validation_layout = QVBoxLayout()

        self.validation_display = QTextEdit()
        self.validation_display.setReadOnly(True)
        self.validation_display.setMaximumHeight(150)
        self.validation_display.setFont(QFont("Courier New", 9))
        self.validation_display.setText("No file validated yet.")
        validation_layout.addWidget(self.validation_display)

        validation_group.setLayout(validation_layout)
        layout.addWidget(validation_group)

        # Action buttons
        button_layout = QHBoxLayout()

        self.validate_btn = QPushButton("Validate Template")
        self.validate_btn.clicked.connect(self._validate_template)
        self.validate_btn.setEnabled(False)
        button_layout.addWidget(self.validate_btn)

        self.import_btn = QPushButton("Import Template")
        self.import_btn.clicked.connect(self._import_template)
        self.import_btn.setEnabled(False)
        button_layout.addWidget(self.import_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self._clear)
        button_layout.addWidget(self.clear_btn)

        button_layout.addStretch()

        layout.addLayout(button_layout)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.status_label)

        layout.addStretch()

    def _browse_file(self):
        """Open file browser to select template"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Template File",
            "",
            "Python Files (*.py);;All Files (*)"
        )

        if file_path:
            self._set_selected_file(file_path)

    def _drag_enter(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].toLocalFile().endswith('.py'):
                event.acceptProposedAction()

    def _drop(self, event: QDropEvent):
        """Handle drop event"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1:
                file_path = urls[0].toLocalFile()
                if file_path.endswith('.py'):
                    self._set_selected_file(file_path)
                    event.acceptProposedAction()
                else:
                    QMessageBox.warning(
                        self,
                        "Invalid File",
                        "Please select a Python (.py) file."
                    )

    def _set_selected_file(self, file_path: str):
        """Set the selected file and update UI"""
        self.selected_file = file_path
        self.file_label.setText(f"Selected: {Path(file_path).name}")
        self.file_label.setStyleSheet("color: #000; font-style: normal;")
        self.validate_btn.setEnabled(True)
        self.import_btn.setEnabled(False)
        self.validation_result = None
        self.validation_display.setText("Click 'Validate Template' to check this file.")
        self.status_label.setText("")

        # Update drop area style
        self.drop_area.setStyleSheet("""
            QLabel {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                padding: 40px;
                background-color: #f0f8f0;
                min-height: 100px;
            }
        """)
        self.drop_area.setText(f"✓ File selected: {Path(file_path).name}\n\nDrop another file to change")

    def _validate_template(self):
        """Validate the selected template file"""
        if not self.selected_file:
            return

        try:
            self.status_label.setText("Validating...")
            self.validate_btn.setEnabled(False)

            # Read file
            with open(self.selected_file, 'r', encoding='utf-8') as f:
                code = f.read()

            # Validate
            self.validation_result = self.template_manager.validate_template(code)

            # Display results
            result_text = []

            if self.validation_result.valid:
                result_text.append("✓ VALIDATION PASSED\n")
                result_text.append("="  * 50)

                if self.validation_result.metadata:
                    result_text.append("\nTemplate Information:")
                    result_text.append(f"  Name: {self.validation_result.metadata.get('name', 'N/A')}")
                    result_text.append(f"  Description: {self.validation_result.metadata.get('description', 'N/A')}")
                    result_text.append(f"  Category: {self.validation_result.metadata.get('category', 'N/A')}")
                    result_text.append(f"  Version: {self.validation_result.metadata.get('version', '1.0.0')}")

                    params = self.validation_result.metadata.get('parameters', {})
                    result_text.append(f"\n  Parameters: {len(params)}")
                    if params:
                        for param_name in list(params.keys())[:5]:  # Show first 5
                            result_text.append(f"    - {param_name}")
                        if len(params) > 5:
                            result_text.append(f"    ... and {len(params) - 5} more")

                if self.validation_result.warnings:
                    result_text.append("\n\nWarnings:")
                    for warning in self.validation_result.warnings:
                        result_text.append(f"  ⚠ {warning}")

                result_text.append("\n\nTemplate is ready to import.")
                self.import_btn.setEnabled(True)
                self.status_label.setText("✓ Validation passed - Ready to import")
                self.status_label.setStyleSheet("color: green; font-weight: bold;")

            else:
                result_text.append("✗ VALIDATION FAILED\n")
                result_text.append("=" * 50)
                result_text.append("\nErrors:")
                for error in self.validation_result.errors:
                    result_text.append(f"  ✗ {error}")

                if self.validation_result.warnings:
                    result_text.append("\nWarnings:")
                    for warning in self.validation_result.warnings:
                        result_text.append(f"  ⚠ {warning}")

                result_text.append("\n\nPlease fix the errors before importing.")
                self.import_btn.setEnabled(False)
                self.status_label.setText("✗ Validation failed - Cannot import")
                self.status_label.setStyleSheet("color: red; font-weight: bold;")

            self.validation_display.setText("\n".join(result_text))

        except Exception as e:
            self.logger.error(f"Validation error: {e}")
            self.validation_display.setText(f"Error during validation:\n\n{str(e)}")
            self.status_label.setText("Error during validation")
            self.status_label.setStyleSheet("color: red;")
            self.import_btn.setEnabled(False)

        finally:
            self.validate_btn.setEnabled(True)

    def _import_template(self):
        """Import the validated template"""
        if not self.selected_file or not self.validation_result or not self.validation_result.valid:
            QMessageBox.warning(
                self,
                "Cannot Import",
                "Please select and validate a template file first."
            )
            return

        try:
            self.status_label.setText("Importing...")
            self.import_btn.setEnabled(False)

            # Get selected category
            category = self.category_combo.currentText()

            # Import template
            imported_path = self.template_manager.import_template(self.selected_file, category)

            if imported_path:
                self.logger.info(f"Template imported: {imported_path}")

                # Show success message
                template_name = self.validation_result.metadata.get('name', 'Template')
                QMessageBox.information(
                    self,
                    "Import Successful",
                    f"Template '{template_name}' has been imported successfully!\n\n"
                    f"Category: {category}\n"
                    f"Location: {Path(imported_path).name}\n\n"
                    "The template is now available in your script library."
                )

                self.status_label.setText("✓ Import successful")
                self.status_label.setStyleSheet("color: green; font-weight: bold;")

                # Emit signal
                self.template_imported.emit(imported_path)

                # Clear for next import
                self._clear()
            else:
                QMessageBox.critical(
                    self,
                    "Import Failed",
                    "Failed to import template. Check logs for details."
                )
                self.status_label.setText("✗ Import failed")
                self.status_label.setStyleSheet("color: red;")

        except Exception as e:
            self.logger.error(f"Import error: {e}")
            QMessageBox.critical(
                self,
                "Import Error",
                f"An error occurred during import:\n\n{str(e)}"
            )
            self.status_label.setText("Error during import")
            self.status_label.setStyleSheet("color: red;")

        finally:
            self.import_btn.setEnabled(self.validation_result and self.validation_result.valid)

    def _clear(self):
        """Clear the uploader"""
        self.selected_file = None
        self.validation_result = None
        self.file_label.setText("No file selected")
        self.file_label.setStyleSheet("color: #666; font-style: italic;")
        self.validation_display.setText("No file validated yet.")
        self.validate_btn.setEnabled(False)
        self.import_btn.setEnabled(False)
        self.status_label.setText("")
        self.category_combo.setCurrentText('Custom')

        # Reset drop area
        self.drop_area.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 5px;
                padding: 40px;
                background-color: #f5f5f5;
                min-height: 100px;
            }
        """)
        self.drop_area.setText("Drag and drop template file here\n\nOR")
