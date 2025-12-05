# Automation Hub

[![CI](https://github.com/criswell82/automation-hub/actions/workflows/ci.yml/badge.svg)](https://github.com/criswell82/automation-hub/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A centralized Python desktop application for automating repetitive tasks in corporate environments.

## Overview

Automation Hub is a PyQt5-based desktop application that provides a unified interface for managing and executing various automation tasks including:

- **Desktop RPA**: UI automation for disconnected systems
- **Excel Automation**: Data manipulation, reporting, and chart generation
- **Outlook Automation**: Email processing, batch sending, and task creation
- **SharePoint Integration**: File management, backups, and search
- **Word Automation**: Document generation and template processing
- **OneNote Integration**: Knowledge management via Microsoft Graph API and COM
- **Asana Integration**: Task management via email, browser automation, and CSV

## Key Features

- **Centralized Dashboard**: Launch and manage all automation scripts from one place
- **Task Scheduling**: Schedule scripts to run at specific times or recurring intervals (APScheduler)
- **AI Workflow Generator**: Create automation scripts from natural language descriptions (Claude AI)
- **Template System**: Browse, import, and create workflow templates
- **Real-time Monitoring**: View script execution status and logs in real-time
- **Configuration Management**: Easily configure script parameters through the GUI
- **Secure Credential Storage**: Windows Credential Manager integration
- **Single Executable**: Deploy as a standalone .exe with no installation required

## Installation

### Prerequisites

- Windows 10 or later
- Python 3.10 or higher (for development)

### For End Users

1. Download the latest `AutomationHub.exe` from the releases page
2. Run the executable - no installation required
3. Configure your settings on first launch

### For Developers

1. Clone this repository:
   ```bash
   git clone https://github.com/criswell82/automation-hub.git
   cd automation-hub
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # On Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Install development dependencies:
   ```bash
   pip install -r requirements-dev.txt
   ```

5. Run the application:
   ```bash
   python src/main.py
   ```

6. Run tests:
   ```bash
   pytest tests/ -v
   ```

## Quick Start

### Running Your First Automation

1. Launch Automation Hub
2. Browse available scripts in the left panel (organized by category)
3. Click on a script to view details
4. Click "Run" to configure parameters and execute
5. View output and logs in the Output tab

### Using the AI Workflow Generator

1. Go to **File > Create New Workflow** (Ctrl+N)
2. Describe what you want to automate in plain English
3. The AI will generate a complete Python workflow script
4. Review, customize, and save the generated script

### Creating a Scheduled Task

1. Select a script from the library
2. Click "Schedule" button
3. Choose schedule type: one-time, interval, or cron
4. Configure parameters and timing
5. The task will run automatically according to schedule

## Project Structure

```
automation_hub/
├── src/
│   ├── main.py                    # Application entry point
│   ├── hub/                       # Central dashboard GUI
│   │   ├── main_window.py         # Main application window
│   │   ├── script_manager.py      # Script discovery and execution
│   │   ├── scheduler.py           # Task scheduling (APScheduler)
│   │   └── ...                    # Dialogs and widgets
│   ├── modules/                   # Automation modules
│   │   ├── base_module.py         # Base class for all modules
│   │   ├── desktop_rpa/           # Desktop RPA engine
│   │   ├── excel_automation/      # Excel automation
│   │   ├── outlook_automation/    # Outlook automation
│   │   ├── sharepoint/            # SharePoint integration
│   │   ├── word_automation/       # Word automation
│   │   ├── onenote/               # OneNote integration
│   │   └── asana/                 # Asana integration
│   ├── core/                      # Infrastructure
│   │   ├── config.py              # Configuration management
│   │   ├── logging_config.py      # Logging framework
│   │   ├── security.py            # Credential management
│   │   ├── error_handler.py       # Error handling
│   │   ├── ai_workflow_generator.py # AI script generation
│   │   └── template_manager.py    # Template management
│   └── utils/                     # Shared utilities
│       ├── report_helpers.py      # Report generation
│       ├── file_helpers.py        # File operations
│       ├── communication_helpers.py # Email workflows
│       ├── onenote_helpers.py     # OneNote workflows
│       └── asana_helpers.py       # Asana workflows
├── tests/                         # Test suite
│   ├── unit/                      # Unit tests
│   └── integration/               # Integration tests
├── docs/                          # Documentation
├── scripts/                       # Build and utility scripts
├── user_scripts/                  # User-created workflows
│   └── templates/                 # Workflow templates
└── resources/                     # Resources (icons, config)
```

## Documentation

- [User Guide](docs/user_guide.md) - How to use Automation Hub
- [Developer Guide](docs/developer_guide.md) - How to create custom automations
- [AI Workflow Generator Guide](docs/AI_Workflow_Generator_Guide.md) - Using AI to create workflows
- [Asana Integration Guide](docs/Asana_Integration_Guide.md) - Asana automation features
- [OneNote Integration Guide](docs/onenote_integration.md) - OneNote automation features
- [Changelog](CHANGELOG.md) - Version history and changes

## Modules

### Desktop RPA
Automate any Windows application through UI interaction:
- Window management (find, activate, position, resize)
- Mouse control (click, drag, scroll)
- Keyboard automation (type text, hotkeys)
- Screenshot capture and image recognition

### Excel Automation
Comprehensive Excel manipulation:
- Read/write data to any cell or range
- Apply formatting (fonts, colors, borders)
- Insert formulas and functions
- Generate charts (bar, line, pie)
- Process multiple workbooks

### Outlook Automation
Email and task management:
- Send emails with attachments
- Read and filter inbox messages
- Process emails automatically
- Create Outlook tasks

### SharePoint Integration
Document management:
- Upload and download files
- Search document libraries
- Manage folder structures
- Backup automation

### Word Automation
Document generation:
- Create documents from templates
- Replace placeholders with data
- Apply formatting
- Generate reports

### OneNote Integration
Knowledge management:
- Create pages and sections
- Insert formatted content
- Organize notebooks
- COM and Graph API support

### Asana Integration
Task management without API:
- Create tasks via email-to-task
- Browser automation for complex operations
- CSV bulk import/export
- Excel-to-Asana synchronization

## Security

- **Credential Management**: Uses Windows Credential Manager for secure storage
- **No Hardcoded Credentials**: All sensitive data stored securely
- **Audit Logging**: Comprehensive logging of all automation activities
- **API Key Management**: Secure storage for API keys (Anthropic, Graph, etc.)

## Building from Source

To create a standalone executable:

```bash
python scripts/build.py
```

The executable will be created in the `dist/` directory.

## Testing

Run the test suite:

```bash
# Run all tests
pytest tests/ -v

# Run with coverage
pytest tests/ -v --cov=src --cov-report=html

# Run specific test file
pytest tests/unit/test_config.py -v
```

## CI/CD

The project uses GitHub Actions for continuous integration:
- Automated testing on Python 3.10, 3.11, and 3.12
- Linting with flake8
- Code coverage reporting
- Tests run on every push and pull request

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run tests (`pytest tests/ -v`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Support

For questions or issues:
- Check the [documentation](docs/)
- Review the [changelog](CHANGELOG.md)
- Open an [issue](https://github.com/criswell82/automation-hub/issues)

## Version

**Current Version:** 1.0.0-beta

See [CHANGELOG.md](CHANGELOG.md) for version history.

## Roadmap

### Completed
- [x] Core infrastructure (config, logging, security, errors)
- [x] Desktop RPA module (window management, input control)
- [x] Excel automation module
- [x] Outlook automation module
- [x] SharePoint integration
- [x] Word automation module
- [x] OneNote integration (COM + Graph API)
- [x] Asana integration (email, browser, CSV)
- [x] Hub GUI with PyQt5
- [x] Script Manager and execution engine
- [x] Task Scheduler (APScheduler)
- [x] AI Workflow Generator (Claude)
- [x] Template management system
- [x] CI/CD pipeline (GitHub Actions)
- [x] Comprehensive test suite (65+ tests)
- [x] Single executable packaging

### In Progress
- [ ] Fix remaining test issues
- [ ] Performance optimization
- [ ] Enhanced error recovery

### Planned
- [ ] Auto-update mechanism
- [ ] Dark/light theme support
- [ ] Plugin system for custom modules
- [ ] Advanced reporting dashboard
- [ ] Multi-language support

---

**Built with:** Python, PyQt5, pywin32, openpyxl, pyautogui, APScheduler, Anthropic

**Tested on:** Windows 10/11, Python 3.10-3.12
