# Automation Hub - Developer Guide

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Development Setup](#development-setup)
3. [Creating Automation Scripts](#creating-automation-scripts)
4. [Module API Reference](#module-api-reference)
5. [Testing](#testing)
6. [Building and Deployment](#building-and-deployment)

## Architecture Overview

Automation Hub follows a modular architecture:

```
┌─────────────────────────────────────────┐
│       Hub (PyQt5 Dashboard)             │
├─────────────────────────────────────────┤
│    Core Infrastructure Layer            │
│  (Config, Logging, Security, etc.)      │
├─────────────────────────────────────────┤
│       Automation Modules                │
│  (RPA, Excel, Outlook, SharePoint...)   │
└─────────────────────────────────────────┘
```

### Key Components

- **Hub**: PyQt5-based GUI for user interaction
- **Core**: Infrastructure services (config, logging, error handling)
- **Modules**: Individual automation capabilities
- **Utils**: Shared utility functions

## Development Setup

### Prerequisites

- Python 3.8+
- Git
- Virtual environment tool

### Initial Setup

```bash
# Clone the repository
git clone <repo-url>
cd automation_hub

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements-dev.txt

# Run tests
pytest

# Run the application
python src/main.py
```

### Project Structure

```
src/
├── main.py                    # Entry point
├── hub/                       # GUI components
│   ├── main_window.py
│   ├── script_manager.py
│   ├── scheduler.py
│   ├── output_viewer.py
│   └── widgets/
├── modules/                   # Automation modules
│   ├── base_module.py        # Base class
│   ├── desktop_rpa/
│   ├── excel_automation/
│   └── ...
├── core/                      # Infrastructure
│   ├── config.py
│   ├── logging_config.py
│   ├── error_handler.py
│   ├── security.py
│   └── powershell_bridge.py
└── utils/                     # Utilities
```

## Creating Automation Scripts

### Basic Script Structure

All automation scripts should inherit from `BaseModule`:

```python
from modules.base_module import BaseModule

class MyAutomation(BaseModule):
    """Description of what this automation does."""

    def __init__(self):
        super().__init__(
            name="My Automation",
            description="Detailed description",
            version="1.0.0"
        )

    def configure(self, **kwargs):
        """
        Configure the automation with parameters.

        Args:
            **kwargs: Configuration parameters

        Returns:
            bool: True if configuration successful
        """
        self.param1 = kwargs.get('param1')
        self.param2 = kwargs.get('param2')
        return self.validate()

    def validate(self):
        """
        Validate configuration before execution.

        Returns:
            bool: True if valid, False otherwise
        """
        if not self.param1:
            self.log_error("param1 is required")
            return False
        return True

    def execute(self):
        """
        Execute the automation.

        Returns:
            dict: Results of execution
        """
        try:
            self.log_info("Starting execution")

            # Your automation logic here
            result = self._do_work()

            self.log_info("Execution completed")
            return {
                'status': 'success',
                'result': result
            }

        except Exception as e:
            self.log_error(f"Execution failed: {str(e)}")
            return {
                'status': 'error',
                'error': str(e)
            }

    def _do_work(self):
        """Internal method for actual work."""
        pass
```

### Using Module Features

#### Logging

```python
self.log_info("Information message")
self.log_warning("Warning message")
self.log_error("Error message")
self.log_debug("Debug message")
```

#### Configuration Access

```python
# Get application config
config = self.get_app_config()

# Get module-specific config
my_setting = self.get_config('my_setting', default='default_value')

# Save config
self.save_config('my_setting', 'new_value')
```

#### Error Handling

```python
try:
    risky_operation()
except Exception as e:
    self.handle_error(e, context="Operation description")
    return self.error_result(str(e))
```

## Module API Reference

### BaseModule

Base class for all automation modules.

#### Methods

**`__init__(name, description, version)`**
Initialize the module with metadata.

**`configure(**kwargs)`**
Configure the module with parameters. Should be overridden.

**`validate()`**
Validate configuration. Should be overridden.

**`execute()`**
Execute the automation. Must be overridden.

**`log_info(message)`, `log_warning(message)`, `log_error(message)`, `log_debug(message)`**
Log messages at different levels.

**`get_config(key, default=None)`**
Retrieve configuration value.

**`save_config(key, value)`**
Save configuration value.

**`handle_error(exception, context='')`**
Standard error handling.

### Desktop RPA Module

```python
from modules.desktop_rpa import WindowManager, InputController

# Window management
wm = WindowManager()
wm.find_window("Window Title")
wm.activate_window()
wm.get_window_rect()

# Input control
ic = InputController()
ic.click(x=100, y=200)
ic.type_text("Hello World")
ic.press_key('enter')
```

### Excel Automation Module

```python
from modules.excel_automation import WorkbookHandler, ChartBuilder

# Workbook operations
wb = WorkbookHandler('path/to/file.xlsx')
wb.read_data('Sheet1', 'A1:D10')
wb.write_data('Sheet1', data, start_cell='A1')
wb.create_chart('Sheet1', chart_type='bar', data_range='A1:D10')
wb.save()
```

### Outlook Automation Module

```python
from modules.outlook_automation import EmailHandler

# Email operations
email = EmailHandler()
email.send_email(
    to='recipient@example.com',
    subject='Subject',
    body='Message body',
    attachments=['file.pdf']
)
email.read_emails(folder='Inbox', unread_only=True)
```

### SharePoint Module

```python
from modules.sharepoint import SharePointClient

# SharePoint operations
sp = SharePointClient(site_url='https://company.sharepoint.com/sites/mysite')
sp.authenticate()
sp.upload_file('local/path.txt', 'SharePoint/Folder/')
sp.download_file('SharePoint/Folder/file.txt', 'local/destination/')
```

## Testing

### Unit Tests

Create tests in `tests/unit/`:

```python
import pytest
from modules.my_module import MyAutomation

def test_configuration():
    automation = MyAutomation()
    result = automation.configure(param1='value1', param2='value2')
    assert result is True

def test_validation():
    automation = MyAutomation()
    automation.configure(param1=None)
    assert automation.validate() is False
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src --cov-report=html

# Run specific test file
pytest tests/unit/test_my_module.py
```

## Building and Deployment

### Creating an Executable

```bash
# Build single executable
python scripts/build.py

# The executable will be in dist/AutomationHub.exe
```

### Build Configuration

Edit `pyinstaller.spec` to customize the build:

- Add additional data files
- Include/exclude modules
- Set application icon
- Configure version information

### Deployment Checklist

- [ ] All tests pass
- [ ] Version number updated
- [ ] CHANGELOG updated
- [ ] Documentation current
- [ ] Build successful
- [ ] Executable tested on clean machine
- [ ] Release notes prepared

## Best Practices

### Code Style

- Follow PEP 8
- Use type hints where appropriate
- Write docstrings for all public methods
- Keep functions focused and small

### Error Handling

- Always use try-except for risky operations
- Log errors with context
- Provide user-friendly error messages
- Don't suppress errors silently

### Security

- Never hardcode credentials
- Use secure credential storage
- Validate all user inputs
- Log security-relevant events

### Performance

- Use lazy loading for heavy modules
- Implement progress indicators for long operations
- Optimize file I/O operations
- Profile performance bottlenecks

## Contributing

### Code Review Process

1. Create feature branch
2. Implement changes with tests
3. Run full test suite
4. Submit for review
5. Address feedback
6. Merge to main

### Commit Messages

Use conventional commits:
- `feat: Add new feature`
- `fix: Fix bug`
- `docs: Update documentation`
- `test: Add tests`
- `refactor: Code refactoring`

---

*For user documentation, see the [User Guide](user_guide.md).*
