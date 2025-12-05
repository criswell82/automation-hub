"""
AI Workflow Generator Dialog

Allows users to describe workflows in natural language and generates Python scripts.
"""

import logging
from typing import Optional, Dict, Any
from pathlib import Path

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QPushButton,
    QLabel, QGroupBox, QLineEdit, QComboBox, QMessageBox,
    QSplitter, QCheckBox, QProgressBar, QTabWidget, QWidget
)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QTextCharFormat, QColor, QSyntaxHighlighter


class PythonSyntaxHighlighter(QSyntaxHighlighter):
    """Simple Python syntax highlighter for code preview"""

    def __init__(self, document):
        super().__init__(document)
        self.setup_formats()

    def setup_formats(self):
        """Setup formatting for different token types"""
        # Keywords
        self.keyword_format = QTextCharFormat()
        self.keyword_format.setForeground(QColor("#0000FF"))
        self.keyword_format.setFontWeight(QFont.Bold)
        self.keywords = [
            'def', 'class', 'return', 'if', 'elif', 'else', 'for', 'while',
            'try', 'except', 'finally', 'import', 'from', 'as', 'with',
            'True', 'False', 'None', 'and', 'or', 'not', 'in', 'is'
        ]

        # Strings
        self.string_format = QTextCharFormat()
        self.string_format.setForeground(QColor("#008000"))

        # Comments
        self.comment_format = QTextCharFormat()
        self.comment_format.setForeground(QColor("#808080"))
        self.comment_format.setFontItalic(True)

    def highlightBlock(self, text):
        """Apply highlighting to a block of text"""
        # Highlight keywords
        for keyword in self.keywords:
            index = text.find(keyword)
            while index >= 0:
                length = len(keyword)
                self.setFormat(index, length, self.keyword_format)
                index = text.find(keyword, index + length)

        # Highlight strings
        for quote in ['"', "'"]:
            index = text.find(quote)
            while index >= 0:
                end_index = text.find(quote, index + 1)
                if end_index == -1:
                    break
                length = end_index - index + 1
                self.setFormat(index, length, self.string_format)
                index = text.find(quote, end_index + 1)

        # Highlight comments
        index = text.find('#')
        if index >= 0:
            self.setFormat(index, len(text) - index, self.comment_format)


class WorkflowGeneratorThread(QThread):
    """Thread for generating workflows using AI"""

    finished = pyqtSignal(dict)  # Emits {success: bool, code: str, error: str}
    progress = pyqtSignal(str)   # Emits progress messages

    def __init__(self, description: str, category: str, use_templates: list):
        super().__init__()
        self.description = description
        self.category = category
        self.use_templates = use_templates
        self.logger = logging.getLogger(__name__)

    def run(self):
        """Generate the workflow"""
        try:
            self.progress.emit("Analyzing your requirements...")

            # Import AI generator
            from core.ai_workflow_generator import AIWorkflowGenerator

            generator = AIWorkflowGenerator()

            self.progress.emit("Generating workflow code...")

            # Generate the workflow
            result = generator.generate_workflow(
                description=self.description,
                category=self.category,
                use_templates=self.use_templates
            )

            self.progress.emit("Workflow generated successfully!")

            self.finished.emit({
                'success': True,
                'code': result.get('code', ''),
                'name': result.get('name', 'Generated Workflow'),
                'description': result.get('description', ''),
                'parameters': result.get('parameters', {}),
                'error': ''
            })

        except Exception as e:
            self.logger.error(f"Workflow generation failed: {e}")
            self.finished.emit({
                'success': False,
                'code': '',
                'error': str(e)
            })


class TemplateCustomizerThread(QThread):
    """Thread for customizing templates using AI"""

    finished = pyqtSignal(dict)  # Emits {success: bool, code: str, error: str}
    progress = pyqtSignal(str)   # Emits progress messages

    def __init__(self, template_code: str, customization_request: str, template_metadata: dict):
        super().__init__()
        self.template_code = template_code
        self.customization_request = customization_request
        self.template_metadata = template_metadata
        self.logger = logging.getLogger(__name__)

    def run(self):
        """Customize the template"""
        try:
            self.progress.emit("Loading template...")

            # Import AI generator
            from core.ai_workflow_generator import AIWorkflowGenerator

            generator = AIWorkflowGenerator()

            self.progress.emit("Customizing template with AI...")

            # Customize the template
            result = generator.customize_template(
                template_code=self.template_code,
                customization_request=self.customization_request,
                template_metadata=self.template_metadata
            )

            self.progress.emit("Template customized successfully!")

            self.finished.emit({
                'success': True,
                'code': result.get('code', ''),
                'name': result.get('name', 'Customized Workflow'),
                'description': result.get('description', ''),
                'parameters': result.get('parameters', {}),
                'error': ''
            })

        except Exception as e:
            self.logger.error(f"Template customization failed: {e}")
            self.finished.emit({
                'success': False,
                'code': '',
                'error': str(e)
            })


class TemplateGeneratorThread(QThread):
    """Thread for generating templates from documents using AI"""

    finished = pyqtSignal(dict)  # Emits {success: bool, code: str, name: str, error: str}
    progress = pyqtSignal(str)   # Emits progress messages

    def __init__(self, analysis: dict, user_instructions: str = ""):
        super().__init__()
        self.analysis = analysis
        self.user_instructions = user_instructions
        self.logger = logging.getLogger(__name__)

    def run(self):
        """Generate template from document analysis"""
        try:
            self.progress.emit("Analyzing document structure...")

            # Import AI generator
            from core.ai_workflow_generator import AIWorkflowGenerator

            generator = AIWorkflowGenerator()

            self.progress.emit("Generating Python template code with AI...")

            # Generate the template
            result = generator.generate_from_document(
                analysis=self.analysis,
                user_instructions=self.user_instructions
            )

            if result.get('success'):
                self.progress.emit("Template generated successfully!")
            else:
                self.progress.emit(f"Generation failed: {result.get('error', 'Unknown error')}")

            self.finished.emit(result)

        except Exception as e:
            self.logger.error(f"Template generation failed: {e}")
            self.finished.emit({
                'success': False,
                'code': '',
                'error': str(e)
            })


class WorkflowGeneratorDialog(QDialog):
    """Dialog for AI-assisted workflow generation"""

    workflow_created = pyqtSignal(str)  # Emits the path to the created workflow file

    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.generated_code = ""
        self.generated_metadata = {}
        self.generator_thread = None
        self.loaded_template = None  # Track if a template is currently loaded
        self.loaded_template_code = None  # Original template code for customization

        # Initialize template manager for loading templates
        from core.template_manager import TemplateManager
        self.template_manager = TemplateManager()

        self.setWindowTitle("Workflow & Template Manager")
        self.resize(1000, 700)

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""
        layout = QVBoxLayout(self)

        # Title
        title_label = QLabel("Workflow & Template Manager")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        # Tab widget
        self.tabs = QTabWidget()

        # Tab 1: AI Generator
        generator_tab = self._create_generator_tab()
        self.tabs.addTab(generator_tab, "Generate with AI")

        # Tab 2: Browse Templates
        browser_tab = self._create_browser_tab()
        self.tabs.addTab(browser_tab, "Browse Templates")

        # Tab 3: Upload Template
        uploader_tab = self._create_uploader_tab()
        self.tabs.addTab(uploader_tab, "Upload Template")

        layout.addWidget(self.tabs)

        # Bottom buttons (common to all tabs)
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.close_btn)

        layout.addLayout(button_layout)

    def _create_generator_tab(self):
        """Create the AI generator tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Description
        desc_label = QLabel(
            "Describe what you want to automate in plain English. "
            "The AI will generate a Python workflow script for you."
        )
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

        # Input section
        input_group = QGroupBox("What do you want to automate?")
        input_layout = QVBoxLayout()

        self.description_input = QTextEdit()
        self.description_input.setPlaceholderText(
            "Example: Create an Excel report from CSV files in a folder, "
            "add charts, and email it to my manager every Monday morning.\n\n"
            "Or: Extract all attachments from unread emails in my inbox "
            "and save them to SharePoint."
        )
        self.description_input.setMinimumHeight(120)
        input_layout.addWidget(self.description_input)

        # Category and options
        options_layout = QHBoxLayout()

        options_layout.addWidget(QLabel("Category:"))
        self.category_combo = QComboBox()
        self.category_combo.addItems([
            "Auto-detect",
            "Report Generation",
            "Email & Communication",
            "File Management",
            "Data Processing",
            "Desktop Automation",
            "Custom"
        ])
        options_layout.addWidget(self.category_combo)

        options_layout.addStretch()

        self.use_templates_check = QCheckBox("Use template examples")
        self.use_templates_check.setChecked(True)
        self.use_templates_check.setToolTip(
            "Include template examples in generation for better results"
        )
        options_layout.addWidget(self.use_templates_check)

        input_layout.addLayout(options_layout)
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        # Template Recommendations section
        self.recommendations_group = QGroupBox("💡 Recommended Templates")
        recommendations_layout = QVBoxLayout()

        self.recommendations_label = QLabel(
            "Start typing your automation description above to see template recommendations..."
        )
        self.recommendations_label.setStyleSheet("color: #666; font-style: italic;")
        recommendations_layout.addWidget(self.recommendations_label)

        # Container for recommendation items
        self.recommendations_container = QWidget()
        self.recommendations_container_layout = QVBoxLayout(self.recommendations_container)
        self.recommendations_container_layout.setContentsMargins(0, 0, 0, 0)
        self.recommendations_container_layout.setSpacing(5)
        recommendations_layout.addWidget(self.recommendations_container)

        self.recommendations_group.setLayout(recommendations_layout)
        self.recommendations_group.setVisible(False)  # Hidden initially
        layout.addWidget(self.recommendations_group)

        # Connect description input to recommendation trigger
        self.description_input.textChanged.connect(self._on_description_changed)

        # Generate button
        generate_btn_layout = QHBoxLayout()
        generate_btn_layout.addStretch()

        self.generate_btn = QPushButton("Generate Workflow")
        self.generate_btn.setMinimumWidth(150)
        self.generate_btn.clicked.connect(self.generate_workflow)
        generate_btn_layout.addWidget(self.generate_btn)

        generate_btn_layout.addStretch()
        layout.addLayout(generate_btn_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Progress message
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        self.progress_label.setStyleSheet("color: #0066CC;")
        layout.addWidget(self.progress_label)

        # Preview section
        preview_group = QGroupBox("Generated Workflow Preview")
        preview_layout = QVBoxLayout()

        # Metadata display
        metadata_layout = QHBoxLayout()
        self.name_display = QLabel("Name: -")
        self.category_display = QLabel("Category: -")
        metadata_layout.addWidget(self.name_display)
        metadata_layout.addWidget(self.category_display)
        metadata_layout.addStretch()
        preview_layout.addLayout(metadata_layout)

        # Code preview
        self.code_preview = QTextEdit()
        self.code_preview.setReadOnly(True)
        self.code_preview.setFont(QFont("Courier New", 9))
        self.code_preview.setMinimumHeight(250)

        # Add syntax highlighting
        self.highlighter = PythonSyntaxHighlighter(self.code_preview.document())

        preview_layout.addWidget(self.code_preview)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Action buttons for generator tab
        button_layout = QHBoxLayout()

        self.save_btn = QPushButton("Save Workflow")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_workflow)
        button_layout.addWidget(self.save_btn)

        self.save_and_run_btn = QPushButton("Save && Run")
        self.save_and_run_btn.setEnabled(False)
        self.save_and_run_btn.clicked.connect(self.save_and_run_workflow)
        button_layout.addWidget(self.save_and_run_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_loaded_template)
        self.clear_btn.setToolTip("Clear loaded template and start fresh")
        button_layout.addWidget(self.clear_btn)

        button_layout.addStretch()

        layout.addLayout(button_layout)

        return tab

    def _create_browser_tab(self):
        """Create the template browser tab"""
        from .template_browser_widget import TemplateBrowserWidget

        browser = TemplateBrowserWidget(self)

        # Connect template selection signal
        browser.template_selected.connect(self._on_template_selected)

        return browser

    def _create_uploader_tab(self):
        """Create the template uploader tab"""
        from .template_uploader_widget import TemplateUploaderWidget

        uploader = TemplateUploaderWidget(self)

        # Connect template import signal to workflow created signal
        uploader.template_imported.connect(self._on_template_imported)

        return uploader

    def _on_template_imported(self, file_path: str):
        """Handle template imported event"""
        self.workflow_created.emit(file_path)

        # Show success and ask if user wants to close
        reply = QMessageBox.question(
            self,
            "Template Imported",
            "Template has been imported successfully!\n\n"
            "Would you like to close this dialog?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.accept()

    def _on_description_changed(self):
        """Handle changes to description input and update recommendations"""
        # Don't show recommendations if a template is already loaded
        if self.loaded_template:
            self.recommendations_group.setVisible(False)
            return

        description = self.description_input.toPlainText().strip()

        # Only show recommendations if description has enough text (at least 10 chars)
        if len(description) < 10:
            self.recommendations_group.setVisible(False)
            return

        # Get template recommendations
        try:
            from core.ai_workflow_generator import AIWorkflowGenerator

            generator = AIWorkflowGenerator()

            # Discover all templates
            templates = self.template_manager.discover_templates()

            # Convert to format expected by recommend_templates
            template_dicts = []
            for template in templates:
                template_dicts.append({
                    'id': template.id,
                    'name': template.name,
                    'description': template.description,
                    'category': template.category,
                    'parameters': template.parameters,
                    'tags': template.tags if hasattr(template, 'tags') else [],
                    'file_path': template.file_path
                })

            # Get recommendations
            recommendations = generator.recommend_templates(description, template_dicts)

            # Display recommendations
            if recommendations:
                self._display_recommendations(recommendations[:3])  # Top 3
                self.recommendations_group.setVisible(True)
            else:
                self.recommendations_group.setVisible(False)

        except Exception as e:
            self.logger.error(f"Failed to get template recommendations: {e}")
            self.recommendations_group.setVisible(False)

    def _display_recommendations(self, recommendations: list):
        """Display template recommendations in the UI"""
        # Clear existing recommendations
        for i in reversed(range(self.recommendations_container_layout.count())):
            widget = self.recommendations_container_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # Update header text
        self.recommendations_label.setText(
            f"Found {len(recommendations)} relevant template(s) for your description:"
        )
        self.recommendations_label.setStyleSheet("color: #000; font-weight: bold;")

        # Add recommendation items
        for rec in recommendations:
            rec_widget = self._create_recommendation_widget(rec)
            self.recommendations_container_layout.addWidget(rec_widget)

    def _create_recommendation_widget(self, recommendation: dict) -> QWidget:
        """Create a widget for a single recommendation"""
        widget = QWidget()
        widget.setStyleSheet("""
            QWidget {
                background-color: #f0f8ff;
                border: 1px solid #b0d4f1;
                border-radius: 5px;
                padding: 8px;
            }
            QWidget:hover {
                background-color: #e0f0ff;
                border-color: #5dade2;
            }
        """)
        widget.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(widget)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)

        # Template name and category
        header_layout = QHBoxLayout()

        name_label = QLabel(f"📋 {recommendation['name']}")
        name_font = QFont()
        name_font.setBold(True)
        name_label.setFont(name_font)
        header_layout.addWidget(name_label)

        category_label = QLabel(f"[{recommendation['category']}]")
        category_label.setStyleSheet("color: #666;")
        header_layout.addWidget(category_label)

        header_layout.addStretch()

        # Score badge
        score = recommendation.get('score', 0)
        score_label = QLabel(f"{score}% match")
        score_label.setStyleSheet("""
            background-color: #4CAF50;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 10px;
        """)
        header_layout.addWidget(score_label)

        layout.addLayout(header_layout)

        # Reason
        reason_label = QLabel(f"💡 {recommendation['reason']}")
        reason_label.setWordWrap(True)
        reason_label.setStyleSheet("color: #444; font-size: 10px;")
        layout.addWidget(reason_label)

        # Click action - load this template
        def on_click(event):
            self._load_recommended_template(recommendation)

        widget.mousePressEvent = on_click

        return widget

    def _load_recommended_template(self, recommendation: dict):
        """Load a recommended template"""
        try:
            # Find the template object by file_path
            templates = self.template_manager.discover_templates()
            template = None

            for t in templates:
                if t.file_path == recommendation['file_path']:
                    template = t
                    break

            if template:
                # Use existing template selection handler
                self._on_template_selected(template)
            else:
                QMessageBox.warning(
                    self,
                    "Template Not Found",
                    f"Could not find template: {recommendation['name']}"
                )

        except Exception as e:
            self.logger.error(f"Failed to load recommended template: {e}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to load template:\n\n{str(e)}"
            )

    def _on_template_selected(self, template):
        """Handle template selection for customization"""
        from core.template_manager import Template

        # Load template code
        full_template = self.template_manager.load_template(template.file_path)

        if not full_template or not full_template.code:
            QMessageBox.warning(
                self,
                "Template Error",
                f"Failed to load template code from:\n{template.file_path}"
            )
            return

        # Build description for AI to customize the template
        description = f"Customize this template: {template.name}\n\n"
        description += f"Description: {template.description}\n\n"

        if template.parameters:
            description += "Parameters:\n"
            for param_name, param_def in list(template.parameters.items())[:5]:  # Show first 5
                if isinstance(param_def, dict):
                    param_desc = param_def.get('description', 'No description')
                    description += f"  - {param_name}: {param_desc}\n"
                else:
                    description += f"  - {param_name}\n"

            if len(template.parameters) > 5:
                description += f"  ... and {len(template.parameters) - 5} more parameters\n"

        description += "\n[Modify the description above to tell AI how to customize this template, or leave as-is to use the template directly]"

        # Populate the generator tab
        self.description_input.setPlainText(description)

        # Set category
        category_index = self.category_combo.findText(template.category)
        if category_index >= 0:
            self.category_combo.setCurrentIndex(category_index)
        else:
            self.category_combo.setCurrentText("Custom")

        # Track loaded template for customization
        self.loaded_template = template
        self.loaded_template_code = full_template.code

        # Display template code in preview
        self.code_preview.setPlainText(full_template.code)
        self.generated_code = full_template.code
        self.generated_metadata = {
            'name': template.name,
            'description': template.description,
            'parameters': template.parameters,
            'category': template.category
        }

        # Update metadata display
        self.name_display.setText(f"Name: {template.name}")
        self.category_display.setText(f"Category: {template.category}")

        # Enable save buttons (template can be used as-is or customized)
        self.save_btn.setEnabled(True)
        self.save_and_run_btn.setEnabled(True)

        # Update Generate button text to indicate customization mode
        self.generate_btn.setText("Customize Template with AI")

        # Hide recommendations when template is loaded
        self.recommendations_group.setVisible(False)

        # Switch to generator tab
        self.tabs.setCurrentIndex(0)

        # Show info message
        QMessageBox.information(
            self,
            "Template Loaded",
            f"Template '{template.name}' has been loaded!\n\n"
            "You can:\n"
            "• Modify the description and click 'Customize Template with AI'\n"
            "• Click 'Save Workflow' to use the template as-is\n"
            "• Click 'Clear' below to start fresh without a template"
        )

    def clear_loaded_template(self):
        """Clear the currently loaded template and reset to normal generation mode"""
        self.loaded_template = None
        self.loaded_template_code = None
        self.generate_btn.setText("Generate Workflow")

        # Clear preview and metadata
        self.code_preview.clear()
        self.generated_code = ""
        self.generated_metadata = {}
        self.name_display.setText("Name: -")
        self.category_display.setText("Category: -")

        # Disable save buttons
        self.save_btn.setEnabled(False)
        self.save_and_run_btn.setEnabled(False)

        # Trigger recommendations refresh if there's description text
        if self.description_input.toPlainText().strip():
            self._on_description_changed()

    def generate_workflow(self):
        """Generate workflow from description or customize loaded template"""
        description = self.description_input.toPlainText().strip()

        if not description:
            QMessageBox.warning(
                self,
                "Missing Description",
                "Please describe what you want to automate."
            )
            return

        # Disable generate button during generation
        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_label.setVisible(True)
        self.progress_label.setText("Initializing...")

        # Check if we're customizing a loaded template
        if self.loaded_template and self.loaded_template_code:
            # Template customization mode
            self.generator_thread = TemplateCustomizerThread(
                self.loaded_template_code,
                description,
                self.generated_metadata
            )
            self.generator_thread.progress.connect(self.on_generation_progress)
            self.generator_thread.finished.connect(self.on_generation_finished)
            self.generator_thread.start()
        else:
            # Normal workflow generation mode
            category = self.category_combo.currentText()
            use_templates = [] if not self.use_templates_check.isChecked() else ['all']

            self.generator_thread = WorkflowGeneratorThread(
                description, category, use_templates
            )
            self.generator_thread.progress.connect(self.on_generation_progress)
            self.generator_thread.finished.connect(self.on_generation_finished)
            self.generator_thread.start()

    def on_generation_progress(self, message: str):
        """Update progress message"""
        self.progress_label.setText(message)

    def on_generation_finished(self, result: dict):
        """Handle generation completion"""
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        if result['success']:
            self.generated_code = result['code']
            self.generated_metadata = {
                'name': result.get('name', 'Generated Workflow'),
                'description': result.get('description', ''),
                'parameters': result.get('parameters', {})
            }

            # Update preview
            self.code_preview.setPlainText(self.generated_code)
            self.name_display.setText(f"Name: {self.generated_metadata['name']}")
            self.category_display.setText(f"Category: {self.category_combo.currentText()}")

            # Enable save buttons
            self.save_btn.setEnabled(True)
            self.save_and_run_btn.setEnabled(True)

            QMessageBox.information(
                self,
                "Success",
                "Workflow generated successfully! Review the code and save when ready."
            )
        else:
            error_msg = result['error']

            # Check if error is related to missing API key
            if 'No API key' in error_msg or 'API key' in error_msg or 'anthropic' in error_msg.lower():
                reply = QMessageBox.question(
                    self,
                    "API Key Required",
                    "AI-powered workflow generation requires an Anthropic API key.\n\n"
                    "Would you like to add your API key now?\n\n"
                    "(The system will use template-based generation as a fallback)",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    # Open settings dialog
                    from .settings_dialog import SettingsDialog
                    settings = SettingsDialog(self)
                    settings.tabs.setCurrentIndex(0)  # Go to API Keys tab
                    if settings.exec_():
                        QMessageBox.information(
                            self,
                            "Settings Saved",
                            "API key saved! You can now try generating your workflow again."
                        )
            else:
                QMessageBox.critical(
                    self,
                    "Generation Failed",
                    f"Failed to generate workflow:\n\n{error_msg}"
                )

    def save_workflow(self):
        """Save the generated workflow to user_scripts/custom/"""
        if not self.generated_code:
            return

        try:
            # Generate filename from name
            safe_name = "".join(c if c.isalnum() or c in (' ', '_') else '_'
                              for c in self.generated_metadata['name'])
            safe_name = safe_name.replace(' ', '_').lower()
            filename = f"{safe_name}.py"

            # Get user_scripts/custom directory
            project_root = Path(__file__).parent.parent.parent
            custom_dir = project_root / "user_scripts" / "custom"
            custom_dir.mkdir(parents=True, exist_ok=True)

            # Save file
            file_path = custom_dir / filename

            # Check if file exists
            if file_path.exists():
                reply = QMessageBox.question(
                    self,
                    "File Exists",
                    f"A workflow named '{filename}' already exists. Overwrite?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.generated_code)

            self.logger.info(f"Workflow saved to: {file_path}")

            QMessageBox.information(
                self,
                "Saved",
                f"Workflow saved successfully!\n\nLocation: {file_path}\n\n"
                "The workflow will appear in your script library after refreshing."
            )

            self.workflow_created.emit(str(file_path))

            return file_path

        except Exception as e:
            self.logger.error(f"Failed to save workflow: {e}")
            QMessageBox.critical(
                self,
                "Save Failed",
                f"Failed to save workflow:\n\n{str(e)}"
            )
            return None

    def save_and_run_workflow(self):
        """Save the workflow and signal to run it immediately"""
        file_path = self.save_workflow()
        if file_path:
            # Accept dialog with saved path
            self.accept()
