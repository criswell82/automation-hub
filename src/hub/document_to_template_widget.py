"""
Document-to-Template Widget

Wizard-style interface for creating Python templates from Word documents.
Users can upload .docx files, analyze structure, and generate templates with AI.
"""
import logging
from pathlib import Path
from typing import Optional, Dict, Any

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QLineEdit, QComboBox, QMessageBox, QProgressBar,
    QStackedWidget, QGroupBox, QScrollArea, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent

from core.document_analyzer import DocumentAnalyzer
from core.template_manager import TemplateManager
from .workflow_generator_dialog import TemplateGeneratorThread


class DragDropArea(QFrame):
    """Drag-and-drop area for file upload."""

    file_dropped = pyqtSignal(str)  # Emits file path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setLineWidth(2)
        self.setMinimumHeight(150)

        # Set up layout
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        # Icon/Label
        label = QLabel("📄 Drag & Drop .docx file here\nor click Browse to select")
        label.setAlignment(Qt.AlignCenter)
        label.setWordWrap(True)
        font = QFont()
        font.setPointSize(11)
        label.setFont(font)
        layout.addWidget(label)

        # Style
        self.setStyleSheet("""
            DragDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #6c757d;
                border-radius: 8px;
            }
            DragDropArea:hover {
                background-color: #e9ecef;
                border-color: #495057;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter."""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].toLocalFile().endswith('.docx'):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    DragDropArea {
                        background-color: #d1ecf1;
                        border: 2px dashed #0c5460;
                        border-radius: 8px;
                    }
                """)

    def dragLeaveEvent(self, event):
        """Handle drag leave."""
        self.setStyleSheet("""
            DragDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #6c757d;
                border-radius: 8px;
            }
            DragDropArea:hover {
                background-color: #e9ecef;
                border-color: #495057;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        """Handle file drop."""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith('.docx'):
                self.file_dropped.emit(file_path)
                event.acceptProposedAction()

        # Reset style
        self.setStyleSheet("""
            DragDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #6c757d;
                border-radius: 8px;
            }
            DragDropArea:hover {
                background-color: #e9ecef;
                border-color: #495057;
            }
        """)


class DocumentToTemplateWidget(QWidget):
    """
    Widget for creating templates from Word documents.

    Workflow:
    1. Upload .docx file (drag-drop or browse)
    2. Analyze document structure
    3. Optionally add customization instructions
    4. Generate Python template with AI
    5. Preview and save to template library
    """

    template_created = pyqtSignal(str)  # Emits path to created template

    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)

        # State
        self.document_path = None
        self.analysis = None
        self.generated_code = None
        self.generated_name = None
        self.generated_description = None
        self.generated_parameters = None

        # Managers
        self.analyzer = DocumentAnalyzer()
        self.template_manager = TemplateManager()
        self.generator_thread = None

        self.setup_ui()

    def setup_ui(self):
        """Set up the user interface."""
        layout = QVBoxLayout(self)

        # Title
        title = QLabel("Create Template from Document")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Description
        desc = QLabel("Upload a Word document and let AI generate a Python template for you.")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Stacked widget for different steps
        self.stack = QStackedWidget()
        layout.addWidget(self.stack)

        # Create all step pages
        self.stack.addWidget(self._create_upload_page())      # Step 0
        self.stack.addWidget(self._create_analysis_page())    # Step 1
        self.stack.addWidget(self._create_customize_page())   # Step 2
        self.stack.addWidget(self._create_generation_page())  # Step 3
        self.stack.addWidget(self._create_preview_page())     # Step 4

        # Navigation buttons
        nav_layout = QHBoxLayout()

        self.back_btn = QPushButton("Back")
        self.back_btn.clicked.connect(self._on_back)
        self.back_btn.setVisible(False)
        nav_layout.addWidget(self.back_btn)

        nav_layout.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self._on_cancel)
        nav_layout.addWidget(self.cancel_btn)

        layout.addLayout(nav_layout)

    def _create_upload_page(self) -> QWidget:
        """Create the document upload page."""
        page = QWidget()
        layout = QVBoxLayout(page)

        # Upload section
        upload_group = QGroupBox("Step 1: Upload Document")
        upload_layout = QVBoxLayout()

        # Drag-drop area
        self.drop_area = DragDropArea()
        self.drop_area.file_dropped.connect(self._on_file_selected)
        upload_layout.addWidget(self.drop_area)

        # Browse button
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._on_browse)
        browse_btn.setMaximumWidth(150)
        browse_layout = QHBoxLayout()
        browse_layout.addStretch()
        browse_layout.addWidget(browse_btn)
        browse_layout.addStretch()
        upload_layout.addLayout(browse_layout)

        # Selected file display
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: #666; font-style: italic;")
        self.file_label.setAlignment(Qt.AlignCenter)
        upload_layout.addWidget(self.file_label)

        upload_group.setLayout(upload_layout)
        layout.addWidget(upload_group)

        # Analyze button
        self.analyze_btn = QPushButton("Analyze Document")
        self.analyze_btn.setMinimumHeight(40)
        self.analyze_btn.setEnabled(False)
        self.analyze_btn.clicked.connect(self._on_analyze)
        analyze_layout = QHBoxLayout()
        analyze_layout.addStretch()
        analyze_layout.addWidget(self.analyze_btn)
        analyze_layout.addStretch()
        layout.addLayout(analyze_layout)

        layout.addStretch()

        return page

    def _create_analysis_page(self) -> QWidget:
        """Create the analysis results page."""
        page = QWidget()
        layout = QVBoxLayout(page)

        # Analysis section
        analysis_group = QGroupBox("Step 2: Analysis Results")
        analysis_layout = QVBoxLayout()

        # Scroll area for results
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(300)

        self.analysis_display = QWidget()
        self.analysis_display_layout = QVBoxLayout(self.analysis_display)
        scroll.setWidget(self.analysis_display)

        analysis_layout.addWidget(scroll)
        analysis_group.setLayout(analysis_layout)
        layout.addWidget(analysis_group)

        # Continue button
        self.continue_btn = QPushButton("Continue to Generate")
        self.continue_btn.setMinimumHeight(40)
        self.continue_btn.clicked.connect(self._on_continue_to_customize)
        continue_layout = QHBoxLayout()
        continue_layout.addStretch()
        continue_layout.addWidget(self.continue_btn)
        continue_layout.addStretch()
        layout.addLayout(continue_layout)

        return page

    def _create_customize_page(self) -> QWidget:
        """Create the customization page."""
        page = QWidget()
        layout = QVBoxLayout(page)

        # Customization section
        custom_group = QGroupBox("Step 3: Customization (Optional)")
        custom_layout = QVBoxLayout()

        label = QLabel("Tell the AI how to customize the template:")
        custom_layout.addWidget(label)

        self.custom_input = QTextEdit()
        self.custom_input.setPlaceholderText(
            "Examples:\n"
            "- Add email validation for email fields\n"
            "- Include error handling for missing files\n"
            "- Add logging for all operations\n"
            "- Make it production-ready with try-catch blocks"
        )
        self.custom_input.setMaximumHeight(120)
        custom_layout.addWidget(self.custom_input)

        custom_group.setLayout(custom_layout)
        layout.addWidget(custom_group)

        # Buttons
        btn_layout = QHBoxLayout()

        skip_btn = QPushButton("Skip")
        skip_btn.clicked.connect(self._on_skip_customize)
        btn_layout.addWidget(skip_btn)

        btn_layout.addStretch()

        generate_btn = QPushButton("Generate Template with AI")
        generate_btn.setMinimumHeight(40)
        generate_btn.clicked.connect(self._on_generate)
        btn_layout.addWidget(generate_btn)

        layout.addLayout(btn_layout)
        layout.addStretch()

        return page

    def _create_generation_page(self) -> QWidget:
        """Create the AI generation page."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignCenter)

        # Generation section
        gen_group = QGroupBox("Step 4: Generating Template")
        gen_layout = QVBoxLayout()
        gen_layout.setAlignment(Qt.AlignCenter)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.setMinimumWidth(400)
        gen_layout.addWidget(self.progress_bar)

        # Status label
        self.progress_label = QLabel("Initializing...")
        self.progress_label.setAlignment(Qt.AlignCenter)
        gen_layout.addWidget(self.progress_label)

        # Cancel generation button
        self.cancel_gen_btn = QPushButton("Cancel Generation")
        self.cancel_gen_btn.clicked.connect(self._on_cancel_generation)
        cancel_layout = QHBoxLayout()
        cancel_layout.addStretch()
        cancel_layout.addWidget(self.cancel_gen_btn)
        cancel_layout.addStretch()
        gen_layout.addLayout(cancel_layout)

        gen_group.setLayout(gen_layout)
        layout.addWidget(gen_group)

        return page

    def _create_preview_page(self) -> QWidget:
        """Create the code preview and save page."""
        page = QWidget()
        layout = QVBoxLayout(page)

        # Preview section
        preview_group = QGroupBox("Step 5: Preview & Save")
        preview_layout = QVBoxLayout()

        # Template info
        info_layout = QHBoxLayout()

        info_layout.addWidget(QLabel("Template Name:"))
        self.name_input = QLineEdit()
        info_layout.addWidget(self.name_input)

        info_layout.addWidget(QLabel("Category:"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["Custom", "Reports", "Email", "Files"])
        info_layout.addWidget(self.category_combo)

        preview_layout.addLayout(info_layout)

        # Code preview
        code_label = QLabel("Generated Code:")
        preview_layout.addWidget(code_label)

        self.code_preview = QTextEdit()
        self.code_preview.setReadOnly(True)
        self.code_preview.setFont(QFont("Courier New", 9))
        self.code_preview.setMinimumHeight(300)
        preview_layout.addWidget(self.code_preview)

        # Action buttons
        action_layout = QHBoxLayout()

        edit_btn = QPushButton("Edit Code")
        edit_btn.clicked.connect(self._on_edit_code)
        action_layout.addWidget(edit_btn)

        regen_btn = QPushButton("Regenerate")
        regen_btn.clicked.connect(self._on_regenerate)
        action_layout.addWidget(regen_btn)

        action_layout.addStretch()

        save_btn = QPushButton("Save to Template Library")
        save_btn.setMinimumHeight(40)
        save_btn.clicked.connect(self._on_save_template)
        action_layout.addWidget(save_btn)

        preview_layout.addLayout(action_layout)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        return page

    # Event Handlers

    def _on_browse(self):
        """Handle browse button click."""
        from PyQt5.QtWidgets import QFileDialog

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Word Document",
            "",
            "Word Documents (*.docx)"
        )

        if file_path:
            self._on_file_selected(file_path)

    def _on_file_selected(self, file_path: str):
        """Handle file selection."""
        self.document_path = file_path
        file_name = Path(file_path).name
        self.file_label.setText(f"Selected: {file_name}")
        self.file_label.setStyleSheet("color: #000; font-weight: bold;")
        self.analyze_btn.setEnabled(True)
        self.logger.info(f"Document selected: {file_path}")

    def _on_analyze(self):
        """Analyze the uploaded document."""
        if not self.document_path:
            return

        try:
            self.analyze_btn.setEnabled(False)
            self.analyze_btn.setText("Analyzing...")

            # Analyze document
            self.analysis = self.analyzer.analyze_word_document(self.document_path)

            # Display results
            self._display_analysis()

            # Move to analysis page
            self.stack.setCurrentIndex(1)
            self.back_btn.setVisible(True)

            self.analyze_btn.setEnabled(True)
            self.analyze_btn.setText("Analyze Document")

        except Exception as e:
            self.logger.error(f"Analysis failed: {e}")
            QMessageBox.critical(
                self,
                "Analysis Failed",
                f"Failed to analyze document:\n\n{str(e)}"
            )
            self.analyze_btn.setEnabled(True)
            self.analyze_btn.setText("Analyze Document")

    def _display_analysis(self):
        """Display analysis results."""
        # Clear existing content
        for i in reversed(range(self.analysis_display_layout.count())):
            widget = self.analysis_display_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        if not self.analysis:
            return

        # Mode and confidence
        mode_layout = QHBoxLayout()
        mode_label = QLabel(f"Template Mode: {self.analysis['mode'].upper()}")
        mode_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        mode_layout.addWidget(mode_label)

        confidence = self.analysis.get('confidence', 0)
        conf_label = QLabel(f"Confidence: {int(confidence * 100)}%")
        conf_label.setStyleSheet("color: #28a745; font-weight: bold;")
        mode_layout.addWidget(conf_label)

        mode_layout.addStretch()
        mode_widget = QWidget()
        mode_widget.setLayout(mode_layout)
        self.analysis_display_layout.addWidget(mode_widget)

        # Template info
        info_text = f"<b>Template Name:</b> {self.analysis['recommended_template_name']}<br>"
        info_text += f"<b>Category:</b> {self.analysis['recommended_category']}"
        info_label = QLabel(info_text)
        self.analysis_display_layout.addWidget(info_label)

        # Structure
        structure = self.analysis['structure']
        struct_text = "<b>Document Structure:</b><ul>"
        struct_text += f"<li>Paragraphs: {structure.get('total_paragraphs', 0)}</li>"
        struct_text += f"<li>Tables: {structure.get('total_tables', 0)}</li>"
        struct_text += f"<li>Headings: {len(structure.get('headings', []))}</li>"
        struct_text += f"<li>Complexity: {structure.get('complexity_score', 0)}/10</li>"
        struct_text += "</ul>"
        struct_label = QLabel(struct_text)
        self.analysis_display_layout.addWidget(struct_label)

        # Placeholders
        placeholders = self.analysis.get('placeholders', [])
        if placeholders:
            ph_text = f"<b>Placeholders Found ({len(placeholders)}):</b><ul>"
            for ph in placeholders[:5]:
                ph_text += f"<li><b>{ph['name']}</b> ({ph['type']}): {ph['description']}</li>"
            if len(placeholders) > 5:
                ph_text += f"<li><i>... and {len(placeholders) - 5} more</i></li>"
            ph_text += "</ul>"
            ph_label = QLabel(ph_text)
            self.analysis_display_layout.addWidget(ph_label)

        self.analysis_display_layout.addStretch()

    def _on_continue_to_customize(self):
        """Move to customization step."""
        self.stack.setCurrentIndex(2)

    def _on_skip_customize(self):
        """Skip customization and generate."""
        self.custom_input.clear()
        self._on_generate()

    def _on_generate(self):
        """Start AI generation."""
        if not self.analysis:
            return

        # Move to generation page
        self.stack.setCurrentIndex(3)
        self.back_btn.setVisible(False)

        # Get customization instructions
        instructions = self.custom_input.toPlainText().strip()

        # Start generation thread
        self.generator_thread = TemplateGeneratorThread(self.analysis, instructions)
        self.generator_thread.progress.connect(self._on_generation_progress)
        self.generator_thread.finished.connect(self._on_generation_finished)
        self.generator_thread.start()

    def _on_generation_progress(self, message: str):
        """Update generation progress."""
        self.progress_label.setText(message)

    def _on_generation_finished(self, result: Dict[str, Any]):
        """Handle generation completion."""
        if result.get('success'):
            # Store generated code
            self.generated_code = result.get('code', '')
            self.generated_name = result.get('name', 'Generated Template')
            self.generated_description = result.get('description', '')
            self.generated_parameters = result.get('parameters', {})

            # Display preview
            self._display_preview()

            # Move to preview page
            self.stack.setCurrentIndex(4)
            self.back_btn.setVisible(True)
        else:
            # Show error
            error_msg = result.get('error', 'Unknown error')
            QMessageBox.critical(
                self,
                "Generation Failed",
                f"Failed to generate template:\n\n{error_msg}\n\n"
                "Please check your API key configuration in Settings."
            )
            # Go back to customize page
            self.stack.setCurrentIndex(2)
            self.back_btn.setVisible(True)

    def _on_cancel_generation(self):
        """Cancel ongoing generation."""
        if self.generator_thread and self.generator_thread.isRunning():
            self.generator_thread.terminate()
            self.generator_thread.wait()

        # Go back to customize page
        self.stack.setCurrentIndex(2)
        self.back_btn.setVisible(True)

    def _display_preview(self):
        """Display generated code preview."""
        self.name_input.setText(self.generated_name)

        # Set category
        category = self.analysis.get('recommended_category', 'Custom')
        index = self.category_combo.findText(category)
        if index >= 0:
            self.category_combo.setCurrentIndex(index)

        # Display code
        self.code_preview.setPlainText(self.generated_code)
        self.code_preview.setReadOnly(True)

    def _on_edit_code(self):
        """Enable code editing."""
        self.code_preview.setReadOnly(False)
        QMessageBox.information(
            self,
            "Edit Mode",
            "Code is now editable. Make your changes and click Save."
        )

    def _on_regenerate(self):
        """Regenerate the template."""
        reply = QMessageBox.question(
            self,
            "Regenerate Template",
            "This will regenerate the template from scratch. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Go back to customize page
            self.stack.setCurrentIndex(2)

    def _on_save_template(self):
        """Save the generated template."""
        if not self.generated_code:
            return

        try:
            # Get template name and category
            template_name = self.name_input.text().strip()
            category = self.category_combo.currentText().lower()

            if not template_name:
                QMessageBox.warning(
                    self,
                    "Invalid Name",
                    "Please enter a template name."
                )
                return

            # Get updated code (in case user edited it)
            code = self.code_preview.toPlainText()

            # Generate filename
            safe_name = "".join(c if c.isalnum() or c in (' ', '_') else '_'
                              for c in template_name)
            safe_name = safe_name.replace(' ', '_').lower()
            filename = f"{safe_name}.py"

            # Save template
            from pathlib import Path
            template_dir = Path(__file__).parent.parent.parent / "user_scripts" / category
            template_dir.mkdir(parents=True, exist_ok=True)

            template_path = template_dir / filename

            # Check if exists
            if template_path.exists():
                reply = QMessageBox.question(
                    self,
                    "File Exists",
                    f"A template named '{filename}' already exists. Overwrite?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return

            # Write template file
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(code)

            # Copy source document to templates/documents/
            if self.document_path and self.analysis.get('mode') == 'fill_in':
                doc_dir = Path(__file__).parent.parent.parent / "user_scripts" / "templates" / "documents"
                doc_dir.mkdir(parents=True, exist_ok=True)

                source_name = Path(self.document_path).name
                dest_path = doc_dir / source_name

                import shutil
                shutil.copy2(self.document_path, dest_path)
                self.logger.info(f"Copied source document to: {dest_path}")

            self.logger.info(f"Template saved: {template_path}")

            # Emit signal
            self.template_created.emit(str(template_path))

            # Show success
            QMessageBox.information(
                self,
                "Template Saved",
                f"Template saved successfully!\n\n"
                f"Location: {template_path}\n\n"
                "The template will appear in your template library."
            )

            # Reset widget
            self._reset()

        except Exception as e:
            self.logger.error(f"Failed to save template: {e}")
            QMessageBox.critical(
                self,
                "Save Failed",
                f"Failed to save template:\n\n{str(e)}"
            )

    def _on_back(self):
        """Go to previous step."""
        current = self.stack.currentIndex()
        if current > 0:
            self.stack.setCurrentIndex(current - 1)
            if current - 1 == 0:
                self.back_btn.setVisible(False)

    def _on_cancel(self):
        """Cancel and reset."""
        reply = QMessageBox.question(
            self,
            "Cancel",
            "Are you sure you want to cancel? All progress will be lost.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self._reset()

    def _reset(self):
        """Reset the widget to initial state."""
        self.document_path = None
        self.analysis = None
        self.generated_code = None
        self.generated_name = None

        self.file_label.setText("No file selected")
        self.file_label.setStyleSheet("color: #666; font-style: italic;")
        self.analyze_btn.setEnabled(False)
        self.custom_input.clear()
        self.code_preview.clear()
        self.name_input.clear()

        self.stack.setCurrentIndex(0)
        self.back_btn.setVisible(False)
