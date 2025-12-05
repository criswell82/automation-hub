"""
Settings Dialog for Automation Hub

Allows users to configure application settings including API keys.
"""

import logging
from typing import Optional

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QPushButton, QLineEdit, QGroupBox, QFormLayout,
    QMessageBox, QDialogButtonBox, QCheckBox, QComboBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont

from core.security import get_security_manager
from core.config import get_config_manager


class APIKeyTestThread(QThread):
    """Thread for testing API key connection"""

    finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, api_key: str):
        super().__init__()
        self.api_key = api_key

    def run(self):
        """Test the API key"""
        try:
            import anthropic

            client = anthropic.Anthropic(api_key=self.api_key)

            # Make a minimal API call to test authentication
            message = client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=10,
                messages=[{"role": "user", "content": "test"}]
            )

            self.finished.emit(True, "API key is valid and working!")

        except ImportError:
            self.finished.emit(
                False,
                "The 'anthropic' package is not installed.\n\n"
                "Install it with: pip install anthropic"
            )
        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg or "authentication" in error_msg.lower():
                self.finished.emit(False, "Invalid API key. Please check and try again.")
            elif "rate" in error_msg.lower():
                self.finished.emit(False, "Rate limit exceeded. The key is valid but you've hit the rate limit.")
            else:
                self.finished.emit(False, f"Connection test failed:\n{error_msg}")


class SettingsDialog(QDialog):
    """Settings dialog with tabs for different configuration areas"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.security_manager = get_security_manager()
        self.config_manager = get_config_manager()

        self.setWindowTitle("Settings")
        self.setMinimumSize(600, 500)

        # Store original values
        self.original_api_key = self._get_current_api_key()

        # Test thread
        self.test_thread = None

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Setup the user interface"""
        layout = QVBoxLayout(self)

        # Create tab widget
        self.tabs = QTabWidget()

        # Add tabs
        self.tabs.addTab(self._create_api_keys_tab(), "API Keys")
        self.tabs.addTab(self._create_general_tab(), "General")
        self.tabs.addTab(self._create_onenote_tab(), "OneNote")
        self.tabs.addTab(self._create_security_tab(), "Security")

        layout.addWidget(self.tabs)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.save_settings)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _create_api_keys_tab(self) -> QWidget:
        """Create the API Keys tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Title and description
        title = QLabel("AI Workflow Generation")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        desc = QLabel(
            "Configure your Anthropic API key to enable AI-powered workflow generation. "
            "The key will be stored securely in Windows Credential Manager."
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

        layout.addSpacing(10)

        # API Key group
        api_group = QGroupBox("Anthropic Claude API Key")
        api_layout = QVBoxLayout()

        # Status label
        self.api_status_label = QLabel()
        self.api_status_label.setStyleSheet("color: #666666;")
        api_layout.addWidget(self.api_status_label)

        # API key input
        input_layout = QHBoxLayout()

        input_label = QLabel("API Key:")
        input_layout.addWidget(input_label)

        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.setPlaceholderText("sk-ant-...")
        self.api_key_input.textChanged.connect(self._on_api_key_changed)
        input_layout.addWidget(self.api_key_input)

        self.show_hide_btn = QPushButton("Show")
        self.show_hide_btn.setMaximumWidth(60)
        self.show_hide_btn.clicked.connect(self._toggle_api_key_visibility)
        input_layout.addWidget(self.show_hide_btn)

        api_layout.addLayout(input_layout)

        # Test button
        test_layout = QHBoxLayout()
        test_layout.addStretch()

        self.test_btn = QPushButton("Test Connection")
        self.test_btn.clicked.connect(self._test_api_key)
        self.test_btn.setEnabled(False)
        test_layout.addWidget(self.test_btn)

        test_layout.addStretch()
        api_layout.addLayout(test_layout)

        # Test result label
        self.test_result_label = QLabel("")
        self.test_result_label.setAlignment(Qt.AlignCenter)
        api_layout.addWidget(self.test_result_label)

        api_group.setLayout(api_layout)
        layout.addWidget(api_group)

        # Help section
        help_group = QGroupBox("How to get an API key")
        help_layout = QVBoxLayout()

        help_text = QLabel(
            "1. Visit <a href='https://console.anthropic.com/'>console.anthropic.com</a><br>"
            "2. Sign up or log in<br>"
            "3. Navigate to 'API Keys' section<br>"
            "4. Create a new API key<br>"
            "5. Copy and paste it above"
        )
        help_text.setOpenExternalLinks(True)
        help_text.setWordWrap(True)
        help_layout.addWidget(help_text)

        help_group.setLayout(help_layout)
        layout.addWidget(help_group)

        layout.addStretch()

        return widget

    def _create_general_tab(self) -> QWidget:
        """Create the General settings tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Model Selection group
        model_group = QGroupBox("AI Model Selection")
        model_layout = QVBoxLayout()

        # Current model display
        self.current_model_label = QLabel()
        self.current_model_label.setStyleSheet("color: #0066CC; font-weight: bold;")
        model_layout.addWidget(self.current_model_label)

        model_layout.addSpacing(5)

        # Model selection dropdown with refresh button
        select_layout = QHBoxLayout()
        select_layout.addWidget(QLabel("Select Model:"))

        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(300)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        select_layout.addWidget(self.model_combo)

        self.refresh_models_btn = QPushButton("Refresh")
        self.refresh_models_btn.setMaximumWidth(80)
        self.refresh_models_btn.clicked.connect(self._refresh_models)
        select_layout.addWidget(self.refresh_models_btn)

        select_layout.addStretch()
        model_layout.addLayout(select_layout)

        # Model info display
        self.model_info_label = QLabel()
        self.model_info_label.setWordWrap(True)
        self.model_info_label.setStyleSheet("color: #666666; font-size: 10px; padding: 5px;")
        model_layout.addWidget(self.model_info_label)

        # Documentation link
        doc_link = QLabel(
            '<a href="https://docs.claude.com/en/docs/about-claude/models">View Model Documentation</a>'
        )
        doc_link.setOpenExternalLinks(True)
        model_layout.addWidget(doc_link)

        model_group.setLayout(model_layout)
        layout.addWidget(model_group)

        # Advanced Settings group
        advanced_group = QGroupBox("Advanced Settings")
        advanced_form = QFormLayout()

        # Max tokens
        max_tokens_label = QLabel("4000")
        advanced_form.addRow("Max Tokens:", max_tokens_label)

        advanced_group.setLayout(advanced_form)
        layout.addWidget(advanced_group)

        layout.addStretch()

        return widget

    def _create_onenote_tab(self) -> QWidget:
        """Create the OneNote settings tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Title and description
        title = QLabel("OneNote Knowledge Management")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        desc = QLabel(
            "Configure default notebooks and sections for OneNote integration. "
            "OneNote serves as your central knowledge repository for project information."
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

        layout.addSpacing(10)

        # Default Notebook/Section group
        defaults_group = QGroupBox("Default Settings")
        defaults_layout = QFormLayout()

        # Default notebook
        self.default_notebook_input = QLineEdit()
        self.default_notebook_input.setPlaceholderText("e.g., 'Project Knowledge'")
        self.default_notebook_input.setToolTip("Default notebook for automation workflows")
        defaults_layout.addRow("Default Notebook:", self.default_notebook_input)

        # Default section
        self.default_section_input = QLineEdit()
        self.default_section_input.setPlaceholderText("e.g., 'Automation'")
        self.default_section_input.setToolTip("Default section within the notebook")
        defaults_layout.addRow("Default Section:", self.default_section_input)

        defaults_group.setLayout(defaults_layout)
        layout.addWidget(defaults_group)

        # Connection test group
        test_group = QGroupBox("Connection Test")
        test_layout = QVBoxLayout()

        test_btn_layout = QHBoxLayout()
        test_btn_layout.addStretch()

        self.onenote_test_btn = QPushButton("Test OneNote Connection")
        self.onenote_test_btn.clicked.connect(self._test_onenote_connection)
        test_btn_layout.addWidget(self.onenote_test_btn)

        test_btn_layout.addStretch()
        test_layout.addLayout(test_btn_layout)

        self.onenote_test_result = QLabel("")
        self.onenote_test_result.setAlignment(Qt.AlignCenter)
        self.onenote_test_result.setWordWrap(True)
        test_layout.addWidget(self.onenote_test_result)

        test_group.setLayout(test_layout)
        layout.addWidget(test_group)

        # Information group
        info_group = QGroupBox("Requirements")
        info_layout = QVBoxLayout()

        info_text = QLabel(
            "• OneNote desktop application must be installed<br>"
            "• OneNote must be signed in with your Microsoft account<br>"
            "• Create notebooks and sections in OneNote before using them in automation<br>"
            "• Uses COM automation - no additional API keys required"
        )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #666666; font-size: 10px;")
        info_layout.addWidget(info_text)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        layout.addStretch()

        return widget

    def _create_security_tab(self) -> QWidget:
        """Create the Security settings tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Credential storage group
        storage_group = QGroupBox("Credential Storage")
        storage_layout = QVBoxLayout()

        self.use_secure_storage_check = QCheckBox(
            "Use Windows Credential Manager for secure storage"
        )
        self.use_secure_storage_check.setChecked(True)
        self.use_secure_storage_check.setEnabled(False)  # Always enabled for security
        storage_layout.addWidget(self.use_secure_storage_check)

        info_label = QLabel(
            "API keys and credentials are stored encrypted in Windows Credential Manager. "
            "This is the most secure method and is always enabled."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666666; font-size: 10px;")
        storage_layout.addWidget(info_label)

        storage_group.setLayout(storage_layout)
        layout.addWidget(storage_group)

        layout.addStretch()

        return widget

    def load_settings(self):
        """Load current settings"""
        # Load API key
        api_key = self._get_current_api_key()
        if api_key:
            self.api_key_input.setText(api_key)
            self._update_api_status(True)
        else:
            self._update_api_status(False)

        # Load model settings
        self._load_models()
        current_model = self.config_manager.get('ai.model', 'claude-sonnet-4-5-20250929')
        self._select_model(current_model)

        # Load OneNote settings
        default_notebook = self.config_manager.get('onenote.default_notebook', '')
        default_section = self.config_manager.get('onenote.default_section', '')
        self.default_notebook_input.setText(default_notebook)
        self.default_section_input.setText(default_section)

    def save_settings(self):
        """Save settings"""
        try:
            api_key = self.api_key_input.text().strip()

            # Save API key if provided
            if api_key:
                # Validate format
                if not api_key.startswith('sk-ant-'):
                    QMessageBox.warning(
                        self,
                        "Invalid API Key",
                        "Anthropic API keys should start with 'sk-ant-'.\n\n"
                        "Please check your key and try again."
                    )
                    return

                # Store in SecurityManager
                self.security_manager.store_credential(
                    target='AnthropicAPIKey',
                    username='default',
                    password=api_key,
                    attributes={'service': 'Anthropic Claude API'}
                )

                self.logger.info("API key saved to secure storage")

            elif self.original_api_key and not api_key:
                # User cleared the API key - delete it
                try:
                    self.security_manager.delete_credential('AnthropicAPIKey')
                    self.logger.info("API key removed from secure storage")
                except Exception:
                    pass

            # Save selected model
            selected_model_data = self.model_combo.currentData()
            if selected_model_data:
                model_id = selected_model_data.get('id')
                if model_id:
                    self.config_manager.set('ai.model', model_id)
                    self.logger.info(f"AI model updated to: {model_id}")

            # Save OneNote settings
            default_notebook = self.default_notebook_input.text().strip()
            default_section = self.default_section_input.text().strip()
            self.config_manager.set('onenote.default_notebook', default_notebook)
            self.config_manager.set('onenote.default_section', default_section)
            if default_notebook or default_section:
                self.logger.info(f"OneNote defaults updated: {default_notebook}/{default_section}")

            # Save config
            self.config_manager.save()

            QMessageBox.information(
                self,
                "Settings Saved",
                "Your settings have been saved successfully!"
            )

            self.accept()

        except Exception as e:
            self.logger.error(f"Failed to save settings: {e}")
            QMessageBox.critical(
                self,
                "Save Failed",
                f"Failed to save settings:\n\n{str(e)}"
            )

    def _get_current_api_key(self) -> Optional[str]:
        """Get the currently stored API key"""
        try:
            cred = self.security_manager.retrieve_credential('AnthropicAPIKey')
            if cred:
                return cred.get('password')
        except Exception:
            pass
        return None

    def _update_api_status(self, has_key: bool):
        """Update the API key status label"""
        if has_key:
            self.api_status_label.setText("✓ API key is configured")
            self.api_status_label.setStyleSheet("color: #008000; font-weight: bold;")
        else:
            self.api_status_label.setText("⚠ No API key configured - AI generation will not be available")
            self.api_status_label.setStyleSheet("color: #CC6600; font-weight: bold;")

    def _on_api_key_changed(self, text: str):
        """Handle API key text change"""
        self.test_btn.setEnabled(bool(text.strip()))
        self.test_result_label.setText("")

    def _toggle_api_key_visibility(self):
        """Toggle API key visibility"""
        if self.api_key_input.echoMode() == QLineEdit.Password:
            self.api_key_input.setEchoMode(QLineEdit.Normal)
            self.show_hide_btn.setText("Hide")
        else:
            self.api_key_input.setEchoMode(QLineEdit.Password)
            self.show_hide_btn.setText("Show")

    def _test_api_key(self):
        """Test the API key connection"""
        api_key = self.api_key_input.text().strip()

        if not api_key:
            return

        # Validate format first
        if not api_key.startswith('sk-ant-'):
            self.test_result_label.setText("⚠ API key should start with 'sk-ant-'")
            self.test_result_label.setStyleSheet("color: #CC6600;")
            return

        # Disable button and show testing message
        self.test_btn.setEnabled(False)
        self.test_result_label.setText("Testing connection...")
        self.test_result_label.setStyleSheet("color: #666666;")

        # Start test thread
        self.test_thread = APIKeyTestThread(api_key)
        self.test_thread.finished.connect(self._on_test_finished)
        self.test_thread.start()

    def _on_test_finished(self, success: bool, message: str):
        """Handle test completion"""
        self.test_btn.setEnabled(True)

        if success:
            self.test_result_label.setText(f"✓ {message}")
            self.test_result_label.setStyleSheet("color: #008000; font-weight: bold;")
        else:
            self.test_result_label.setText(f"✗ {message}")
            self.test_result_label.setStyleSheet("color: #CC0000;")

    def _load_models(self):
        """Load available models into the dropdown"""
        from core.ai_workflow_generator import AIWorkflowGenerator

        # Get available models
        models = AIWorkflowGenerator.get_available_models()

        # Clear and populate dropdown
        self.model_combo.clear()

        for model in models:
            display_text = f"{model['name']} - {model['description']}"
            self.model_combo.addItem(display_text, model)

        self.logger.debug(f"Loaded {len(models)} models into dropdown")

    def _select_model(self, model_id: str):
        """Select a model in the dropdown by ID"""
        for i in range(self.model_combo.count()):
            model_data = self.model_combo.itemData(i)
            if model_data and model_data.get('id') == model_id:
                self.model_combo.setCurrentIndex(i)
                self._update_model_display(model_data)
                break

    def _refresh_models(self):
        """Refresh the models list"""
        self.refresh_models_btn.setEnabled(False)
        self.refresh_models_btn.setText("...")

        # Get API key if available
        api_key = self._get_current_api_key()

        try:
            from core.ai_workflow_generator import AIWorkflowGenerator

            # Reload models (with API key if available)
            models = AIWorkflowGenerator.get_available_models(api_key)

            # Remember current selection
            current_data = self.model_combo.currentData()
            current_id = current_data.get('id') if current_data else None

            # Repopulate
            self.model_combo.clear()
            for model in models:
                display_text = f"{model['name']} - {model['description']}"
                self.model_combo.addItem(display_text, model)

            # Restore selection
            if current_id:
                self._select_model(current_id)

            self.logger.info("Models refreshed successfully")

        except Exception as e:
            self.logger.error(f"Failed to refresh models: {e}")
            QMessageBox.warning(
                self,
                "Refresh Failed",
                f"Failed to refresh models:\n\n{str(e)}"
            )

        finally:
            self.refresh_models_btn.setEnabled(True)
            self.refresh_models_btn.setText("Refresh")

    def _on_model_changed(self, index: int):
        """Handle model selection change"""
        model_data = self.model_combo.itemData(index)
        if model_data:
            self._update_model_display(model_data)

    def _update_model_display(self, model_data: dict):
        """Update the model information display"""
        self.current_model_label.setText(f"Current: {model_data['name']}")

        info_text = f"<b>{model_data['description']}</b><br>"
        info_text += f"Speed: {model_data['speed']} | "
        info_text += f"Capability: {model_data['capability']}"

        self.model_info_label.setText(info_text)

    def _test_onenote_connection(self):
        """Test OneNote connection"""
        self.onenote_test_btn.setEnabled(False)
        self.onenote_test_result.setText("Testing connection...")
        self.onenote_test_result.setStyleSheet("color: #666666;")

        try:
            from modules.onenote import OneNoteCOMClient

            # Try to connect to OneNote
            client = OneNoteCOMClient()
            client.connect()

            # Get notebooks to verify connection
            notebooks = client.get_notebooks()

            # Success!
            self.onenote_test_result.setText(
                f"✓ Connected successfully! Found {len(notebooks)} notebook(s)."
            )
            self.onenote_test_result.setStyleSheet("color: #008000; font-weight: bold;")

            self.logger.info(f"OneNote connection test successful: {len(notebooks)} notebooks found")

        except ImportError as e:
            self.onenote_test_result.setText(
                "✗ pywin32 package not installed. Run: pip install pywin32"
            )
            self.onenote_test_result.setStyleSheet("color: #CC0000;")
            self.logger.error(f"OneNote connection test failed: {e}")

        except Exception as e:
            error_msg = str(e)
            if "OneNote" in error_msg:
                self.onenote_test_result.setText(
                    "✗ OneNote desktop application not found. Please install and sign in to OneNote."
                )
            else:
                self.onenote_test_result.setText(
                    f"✗ Connection failed: {error_msg}"
                )
            self.onenote_test_result.setStyleSheet("color: #CC0000;")
            self.logger.error(f"OneNote connection test failed: {e}")

        finally:
            self.onenote_test_btn.setEnabled(True)
