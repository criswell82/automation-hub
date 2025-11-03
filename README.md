# Automation Hub

A centralized Python desktop application for automating repetitive tasks in corporate environments.

## Overview

Automation Hub is a PyQt5-based desktop application that provides a unified interface for managing and executing various automation tasks including:

- **Desktop RPA**: UI automation for disconnected systems
- **Excel Automation**: Data manipulation, reporting, and chart generation
- **Outlook Automation**: Email processing, batch sending, and task creation
- **SharePoint Integration**: File management, backups, and search
- **Word Automation**: Document generation and formatting
- **OneNote Integration**: Knowledge management via Microsoft Graph API

## Key Features

- **Centralized Dashboard**: Launch and manage all automation scripts from one place
- **Task Scheduling**: Schedule scripts to run at specific times or recurring intervals
- **Real-time Monitoring**: View script execution status and logs in real-time
- **Configuration Management**: Easily configure script parameters through the GUI
- **Self-Updating**: Automatic update mechanism to stay current
- **Single Executable**: Deploy as a standalone .exe with no installation required

## Installation

### Prerequisites

- Windows 10 or later
- Python 3.8 or higher (for development)

### For End Users

1. Download the latest `AutomationHub.exe` from the releases page
2. Run the executable - no installation required
3. Configure your settings on first launch

### For Developers

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd automation_hub
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

4. Install development dependencies (optional):
   ```bash
   pip install -r requirements-dev.txt
   ```

5. Run the application:
   ```bash
   python src/main.py
   ```

## Quick Start

### Running Your First Automation

1. Launch Automation Hub
2. Browse available scripts in the dashboard
3. Click on a script to configure parameters
4. Click "Run" to execute immediately or "Schedule" to set up recurring execution
5. View output and logs in the Output Viewer panel

### Creating a New Automation Script

See the [Developer Guide](docs/developer_guide.md) for detailed instructions on creating custom automation scripts.

## Project Structure

```
automation_hub/
├── src/
│   ├── main.py                    # Application entry point
│   ├── hub/                       # Central dashboard GUI
│   ├── modules/                   # Automation modules
│   │   ├── desktop_rpa/          # Desktop RPA engine
│   │   ├── excel_automation/     # Excel automation
│   │   ├── outlook_automation/   # Outlook automation
│   │   ├── sharepoint/           # SharePoint integration
│   │   ├── word_automation/      # Word automation
│   │   └── onenote/              # OneNote integration
│   ├── core/                      # Infrastructure (config, logging, etc.)
│   └── utils/                     # Shared utilities
├── tests/                         # Unit and integration tests
├── docs/                          # Documentation
├── scripts/                       # Build and utility scripts
└── resources/                     # Resources (icons, templates, config)
```

## Documentation

- [User Guide](docs/user_guide.md) - How to use Automation Hub
- [Developer Guide](docs/developer_guide.md) - How to create custom automations
- [Module Reference](docs/module_reference/) - API documentation for each module
- [Examples](docs/examples/) - Sample automation scripts

## Use Cases

### Data Transfer Automation
Copy data from locked-down internal tools to Excel for analysis and reporting.

### Weekly Reporting
Automatically generate weekly metrics reports with charts and pivot tables.

### Email Processing
Process incoming emails, extract attachments, and create tasks automatically.

### SharePoint Backups
Automatically backup research notes and documents to SharePoint on a schedule.

### Document Generation
Generate standardized meeting minutes and status reports from templates.

## Security

- **Credential Management**: Uses Windows Credential Manager for secure storage
- **No Hardcoded Credentials**: All sensitive data is stored securely
- **Audit Logging**: Comprehensive logging of all automation activities
- **Corporate Compliance**: Designed to work within IT security policies

## Building from Source

To create a standalone executable:

```bash
python scripts/build.py
```

The executable will be created in the `dist/` directory.

## Contributing

This is a corporate internal tool. For feature requests or bug reports, please contact the development team.

## License

Internal use only - see LICENSE file for details.

## Support

For questions or issues:
- Check the [documentation](docs/)
- Review [example scripts](docs/examples/)
- Contact the development team

## Version

Current Version: 1.0.0-alpha

## Roadmap

- [x] Project setup and infrastructure
- [ ] Desktop RPA module
- [ ] Hub GUI interface
- [ ] Excel automation
- [ ] Outlook automation
- [ ] SharePoint integration
- [ ] Word automation
- [ ] OneNote integration
- [ ] Single executable packaging
- [ ] Self-update mechanism
- [ ] Version 1.0 release
