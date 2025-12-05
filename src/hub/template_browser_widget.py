"""
Template Browser Widget

UI for browsing, searching, and selecting workflow templates.
"""

import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QLineEdit, QSplitter, QGroupBox, QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont

from core.template_manager import TemplateManager, Template


class TemplateBrowserWidget(QWidget):
    """Widget for browsing and selecting templates"""

    template_selected = pyqtSignal(Template)  # Emits selected template for use
    template_exported = pyqtSignal(str)  # Emits path to exported template

    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.template_manager = TemplateManager()
        self.templates = []
        self.selected_template = None

        self._create_ui()
        self._load_templates()

    def _create_ui(self):
        """Create the browser UI"""
        layout = QVBoxLayout(self)

        # Instructions
        instructions = QLabel(
            "Browse available workflow templates. Select a template to view details, "
            "or click 'Use Template' to customize it with AI."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        # Search bar
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Type to filter templates...")
        self.search_input.textChanged.connect(self._filter_templates)
        search_layout.addWidget(self.search_input)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self._load_templates)
        search_layout.addWidget(self.refresh_btn)

        layout.addLayout(search_layout)

        # Splitter for tree and preview
        splitter = QSplitter(Qt.Horizontal)

        # Left: Template tree
        tree_widget = QWidget()
        tree_layout = QVBoxLayout(tree_widget)
        tree_layout.setContentsMargins(0, 0, 0, 0)

        tree_label = QLabel("Template Library")
        tree_label_font = QFont()
        tree_label_font.setBold(True)
        tree_label.setFont(tree_label_font)
        tree_layout.addWidget(tree_label)

        self.template_tree = QTreeWidget()
        self.template_tree.setHeaderLabel("Templates")
        self.template_tree.itemClicked.connect(self._on_template_clicked)
        self.template_tree.itemDoubleClicked.connect(self._on_template_double_clicked)
        tree_layout.addWidget(self.template_tree)

        splitter.addWidget(tree_widget)

        # Right: Preview panel
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        preview_label = QLabel("Template Preview")
        preview_label_font = QFont()
        preview_label_font.setBold(True)
        preview_label.setFont(preview_label_font)
        preview_layout.addWidget(preview_label)

        self.preview_display = QTextEdit()
        self.preview_display.setReadOnly(True)
        self.preview_display.setFont(QFont("Courier New", 9))
        self.preview_display.setText("Select a template to view details.")
        preview_layout.addWidget(self.preview_display)

        # Action buttons
        button_layout = QHBoxLayout()

        self.use_btn = QPushButton("Use Template")
        self.use_btn.setEnabled(False)
        self.use_btn.clicked.connect(self._use_template)
        self.use_btn.setToolTip("Load this template for customization with AI")
        button_layout.addWidget(self.use_btn)

        self.export_btn = QPushButton("Export")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export_template)
        self.export_btn.setToolTip("Export template to a file")
        button_layout.addWidget(self.export_btn)

        self.delete_btn = QPushButton("Delete")
        self.delete_btn.setEnabled(False)
        self.delete_btn.clicked.connect(self._delete_template)
        self.delete_btn.setToolTip("Delete custom template (system templates cannot be deleted)")
        button_layout.addWidget(self.delete_btn)

        button_layout.addStretch()

        preview_layout.addLayout(button_layout)

        splitter.addWidget(preview_widget)

        # Set initial sizes
        splitter.setSizes([300, 500])

        layout.addWidget(splitter)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.status_label)

    def _load_templates(self):
        """Load all templates"""
        try:
            self.status_label.setText("Loading templates...")
            self.templates = self.template_manager.discover_templates()
            self._populate_tree()
            self.status_label.setText(f"Loaded {len(self.templates)} templates")
            self.logger.info(f"Loaded {len(self.templates)} templates")
        except Exception as e:
            self.logger.error(f"Failed to load templates: {e}")
            self.status_label.setText("Error loading templates")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to load templates:\n\n{str(e)}"
            )

    def _populate_tree(self):
        """Populate the template tree"""
        self.template_tree.clear()

        # Group templates by category
        templates_by_category = {}
        for template in self.templates:
            category = template.category
            if category not in templates_by_category:
                templates_by_category[category] = []
            templates_by_category[category].append(template)

        # Create tree items
        for category in sorted(templates_by_category.keys()):
            category_templates = templates_by_category[category]

            # Create category item
            category_item = QTreeWidgetItem()
            category_item.setText(0, f"{category} ({len(category_templates)})")
            category_item.setData(0, Qt.UserRole, None)  # Mark as category

            # Make category bold
            font = category_item.font(0)
            font.setBold(True)
            category_item.setFont(0, font)

            # Add templates under category
            for template in sorted(category_templates, key=lambda t: t.name):
                template_item = QTreeWidgetItem(category_item)
                template_item.setText(0, template.name)
                template_item.setData(0, Qt.UserRole, template)  # Store template object
                template_item.setToolTip(0, template.description)

                # Mark user templates with icon or style
                if not template.is_system:
                    template_item.setText(0, f"📁 {template.name}")

            self.template_tree.addTopLevelItem(category_item)

        # Expand all categories
        self.template_tree.expandAll()

    def _filter_templates(self, search_text: str):
        """Filter templates based on search text"""
        search_text = search_text.lower()

        for i in range(self.template_tree.topLevelItemCount()):
            category_item = self.template_tree.topLevelItem(i)
            category_visible = False

            for j in range(category_item.childCount()):
                template_item = category_item.child(j)
                template = template_item.data(0, Qt.UserRole)

                if template:
                    # Check if search matches name, description, or category
                    matches = (
                        search_text in template.name.lower() or
                        search_text in template.description.lower() or
                        search_text in template.category.lower()
                    )
                    template_item.setHidden(not matches)
                    if matches:
                        category_visible = True
                else:
                    template_item.setHidden(False)

            # Hide category if no children match
            category_item.setHidden(not category_visible and bool(search_text))

    def _on_template_clicked(self, item: QTreeWidgetItem):
        """Handle template selection"""
        template = item.data(0, Qt.UserRole)

        if template is None:
            # Category was clicked
            self.selected_template = None
            self.use_btn.setEnabled(False)
            self.export_btn.setEnabled(False)
            self.delete_btn.setEnabled(False)
            self.preview_display.setText("Select a template to view details.")
            return

        self.selected_template = template
        self._update_preview(template)

        # Enable buttons
        self.use_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.delete_btn.setEnabled(not template.is_system)

        self.status_label.setText(f"Selected: {template.name}")

    def _on_template_double_clicked(self, item: QTreeWidgetItem):
        """Handle double-click on template (same as Use Template)"""
        template = item.data(0, Qt.UserRole)
        if template:
            self._use_template()

    def _update_preview(self, template: Template):
        """Update the preview panel with template details"""
        preview_lines = []

        # Header
        preview_lines.append("=" * 60)
        preview_lines.append(f"TEMPLATE: {template.name}")
        preview_lines.append("=" * 60)
        preview_lines.append("")

        # Metadata
        preview_lines.append("📋 INFORMATION")
        preview_lines.append("-" * 60)
        preview_lines.append(f"Name:        {template.name}")
        preview_lines.append(f"Category:    {template.category}")
        preview_lines.append(f"Version:     {template.version}")
        preview_lines.append(f"Author:      {template.author}")
        preview_lines.append(f"Type:        {'System Template' if template.is_system else 'User Template'}")
        preview_lines.append("")

        # Description
        preview_lines.append("📝 DESCRIPTION")
        preview_lines.append("-" * 60)
        preview_lines.append(template.description)
        preview_lines.append("")

        # Parameters
        if template.parameters:
            preview_lines.append(f"⚙️ PARAMETERS ({len(template.parameters)})")
            preview_lines.append("-" * 60)

            for param_name, param_def in template.parameters.items():
                if isinstance(param_def, dict):
                    param_type = param_def.get('type', 'string')
                    param_desc = param_def.get('description', 'No description')
                    param_required = param_def.get('required', False)
                    param_default = param_def.get('default', None)

                    req_marker = "✓" if param_required else "○"
                    preview_lines.append(f"{req_marker} {param_name} ({param_type})")
                    preview_lines.append(f"    {param_desc}")

                    if param_default is not None:
                        preview_lines.append(f"    Default: {param_default}")

                    if param_type == 'choice' and 'choices' in param_def:
                        choices_str = ", ".join(str(c) for c in param_def['choices'])
                        preview_lines.append(f"    Choices: {choices_str}")

                    preview_lines.append("")
                else:
                    preview_lines.append(f"○ {param_name}")
                    preview_lines.append("")
        else:
            preview_lines.append("⚙️ PARAMETERS")
            preview_lines.append("-" * 60)
            preview_lines.append("No parameters defined")
            preview_lines.append("")

        # File location
        preview_lines.append("📁 FILE LOCATION")
        preview_lines.append("-" * 60)
        preview_lines.append(template.file_path)
        preview_lines.append("")

        # Usage hint
        preview_lines.append("💡 USAGE")
        preview_lines.append("-" * 60)
        preview_lines.append("• Click 'Use Template' to customize this template with AI")
        preview_lines.append("• Click 'Export' to save a copy to your local filesystem")
        if not template.is_system:
            preview_lines.append("• Click 'Delete' to remove this custom template")

        self.preview_display.setText("\n".join(preview_lines))

    def _use_template(self):
        """Emit signal to use the selected template"""
        if not self.selected_template:
            return

        self.logger.info(f"Using template: {self.selected_template.name}")
        self.template_selected.emit(self.selected_template)

    def _export_template(self):
        """Export the selected template to a file"""
        if not self.selected_template:
            return

        try:
            # Generate default filename
            safe_name = "".join(c if c.isalnum() or c in (' ', '_') else '_'
                              for c in self.selected_template.name)
            safe_name = safe_name.replace(' ', '_').lower()
            default_filename = f"{safe_name}.py"

            # Show save dialog
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Template",
                default_filename,
                "Python Files (*.py);;All Files (*)"
            )

            if file_path:
                # Export template
                success = self.template_manager.export_template(
                    self.selected_template.id,
                    file_path
                )

                if success:
                    QMessageBox.information(
                        self,
                        "Export Successful",
                        f"Template exported successfully to:\n\n{file_path}"
                    )
                    self.template_exported.emit(file_path)
                    self.logger.info(f"Template exported to: {file_path}")
                else:
                    QMessageBox.critical(
                        self,
                        "Export Failed",
                        "Failed to export template. Check logs for details."
                    )

        except Exception as e:
            self.logger.error(f"Export error: {e}")
            QMessageBox.critical(
                self,
                "Export Error",
                f"An error occurred during export:\n\n{str(e)}"
            )

    def _delete_template(self):
        """Delete the selected custom template"""
        if not self.selected_template:
            return

        # Only allow deleting custom templates
        if self.selected_template.is_system:
            QMessageBox.warning(
                self,
                "Cannot Delete",
                "System templates cannot be deleted."
            )
            return

        # Confirm deletion
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete this template?\n\n"
            f"Template: {self.selected_template.name}\n\n"
            f"This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                success = self.template_manager.delete_template(self.selected_template.id)

                if success:
                    QMessageBox.information(
                        self,
                        "Deleted",
                        f"Template '{self.selected_template.name}' has been deleted."
                    )
                    self.logger.info(f"Template deleted: {self.selected_template.name}")

                    # Reload templates
                    self._load_templates()

                    # Clear preview
                    self.selected_template = None
                    self.preview_display.setText("Select a template to view details.")
                    self.use_btn.setEnabled(False)
                    self.export_btn.setEnabled(False)
                    self.delete_btn.setEnabled(False)
                else:
                    QMessageBox.critical(
                        self,
                        "Deletion Failed",
                        "Failed to delete template. Check logs for details."
                    )

            except Exception as e:
                self.logger.error(f"Delete error: {e}")
                QMessageBox.critical(
                    self,
                    "Delete Error",
                    f"An error occurred during deletion:\n\n{str(e)}"
                )

    def refresh(self):
        """Public method to refresh template list"""
        self._load_templates()
