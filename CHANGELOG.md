# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Nothing yet

---

## [1.0.0-beta] - 2025-12-05

### Added
- **GitHub Actions CI/CD Pipeline**
  - Automated testing on Python 3.10, 3.11, and 3.12
  - Windows-based test runners
  - Linting with flake8
  - Code coverage reporting with Codecov
  - Automatic CI runs on push and pull requests

- **Comprehensive Test Suite**
  - 65 pytest tests covering core functionality
  - Unit tests for ConfigManager, SecurityManager, TemplateManager
  - Integration tests for Excel automation and Document Analyzer
  - pytest-qt support for GUI testing

- **Asana Integration Module** (`src/modules/asana/`)
  - Email-to-task creation (no API required)
  - Browser automation for complex operations
  - CSV import/export for bulk operations
  - Excel-to-Asana synchronization
  - AsanaHelper workflow utilities

- **AI Workflow Generator** (`src/core/ai_workflow_generator.py`)
  - Claude API integration for intelligent script generation
  - Natural language to Python workflow conversion
  - Template-based fallback when AI unavailable
  - Document-to-template generation
  - Template customization and recommendations

- **Template Management System**
  - Template browser with search and filtering
  - Template uploader for importing custom templates
  - Document-to-template converter
  - WORKFLOW_META standard for template metadata

- **Enhanced OneNote Integration**
  - COM client for direct OneNote access
  - Microsoft Graph API support
  - Content formatting (tables, links, images)
  - Notebook/section/page navigation

- **Workflow Helper Modules** (refactored from monolithic file)
  - `src/utils/report_helpers.py` - ReportBuilder, data aggregation
  - `src/utils/onenote_helpers.py` - OneNoteHelper class
  - `src/utils/communication_helpers.py` - EmailHelper, daily reports
  - `src/utils/file_helpers.py` - FileOrganizer, SharePointHelper
  - `src/utils/asana_helpers.py` - AsanaHelper, roadmap creation

### Changed
- **Code Quality Improvements**
  - Added comprehensive type hints to 8 priority modules
  - Fixed all bare `except:` blocks with `except Exception:`
  - Improved module organization and separation of concerns
  - Enhanced docstrings across codebase

- **Build System Improvements**
  - Added resources directory structure
  - Fixed Unicode handling in build script
  - Improved PyInstaller configuration

- **Test Improvements**
  - Fixed Excel format_cells test assertions
  - Fixed Python < 3.12 f-string compatibility
  - Improved CI pipeline reliability

### Fixed
- TypeError in `get_recent_logs` when handling Path objects
- Python 3.10 compatibility issues
- Windows-specific path handling in tests
- Excel fill color test assertions

---

## [1.0.0-alpha] - 2025-11-03

### Added
- **Core Infrastructure**
  - Configuration Management (`src/core/config.py`)
    - JSON/YAML configuration system
    - Nested key access with dot notation
    - Import/export functionality
    - Auto-generated defaults

  - Logging Framework (`src/core/logging_config.py`)
    - Rotating file handlers (10MB, 5 backups)
    - Multiple log levels and formatters
    - Module-specific log files
    - Log archiving and cleanup

  - Error Handling (`src/core/error_handler.py`)
    - Centralized exception management
    - Error categorization and severity
    - User-friendly error messages
    - Error statistics tracking

  - Security Manager (`src/core/security.py`)
    - Windows Credential Manager integration
    - Secure credential storage/retrieval
    - Fallback encrypted file storage
    - API key generation

  - PowerShell Bridge (`src/core/powershell_bridge.py`)
    - Python-to-PowerShell command execution
    - Script execution with parameters
    - Environment variable management

- **Desktop RPA Module** (`src/modules/desktop_rpa/`)
  - Window Manager
    - Find windows by title, regex, class, or process
    - Window state management (activate, minimize, maximize)
    - Position and resize operations
    - Robust retry and timeout logic

  - Input Controller
    - Mouse control (click, move, drag, scroll)
    - Keyboard automation (type, hotkeys)
    - Screenshot capture
    - Image recognition and clicking

- **Excel Automation Module** (`src/modules/excel_automation/`)
  - Workbook creation and manipulation
  - Cell formatting (fonts, colors, borders)
  - Formula insertion
  - Chart generation (bar, line, pie)
  - Auto-size columns

- **Outlook Automation Module** (`src/modules/outlook_automation/`)
  - Email sending with attachments
  - Inbox reading and filtering
  - Task creation from emails
  - COM interface integration

- **SharePoint Integration** (`src/modules/sharepoint/`)
  - File upload/download
  - Document library management
  - Office365-REST-Python-Client integration
  - Search capabilities

- **Word Automation Module** (`src/modules/word_automation/`)
  - Document creation and editing
  - Template-based generation
  - Placeholder replacement
  - Text formatting with python-docx

- **OneNote Integration** (`src/modules/onenote/`)
  - Microsoft Graph API connection
  - Page creation and management
  - Content insertion

- **Hub GUI** (`src/hub/`)
  - Main Window with PyQt5
    - Script library browser with categories
    - Dashboard with statistics
    - Real-time output/log viewer
    - Scheduled tasks panel
    - Menu bar with keyboard shortcuts
    - Status bar

  - Script Manager
    - Dynamic script discovery
    - Script execution engine
    - Parameter configuration

  - Task Scheduler
    - APScheduler integration
    - One-time and recurring schedules
    - Cron-style scheduling

  - Script Dialog
    - Parameter input forms
    - Dry run mode
    - Execution confirmation

- **Build & Deployment**
  - PyInstaller configuration for single executable
  - Build script with dependency verification
  - Resource bundling

- **Documentation**
  - README.md with installation instructions
  - User Guide (`docs/user_guide.md`)
  - Developer Guide (`docs/developer_guide.md`)
  - AI Workflow Generator Guide (`docs/AI_Workflow_Generator_Guide.md`)
  - Asana Integration Guide (`docs/Asana_Integration_Guide.md`)
  - OneNote Integration Guide (`docs/onenote_integration.md`)

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| 1.0.0-beta | 2025-12-05 | CI/CD, test suite, Asana integration, AI generator, code quality |
| 1.0.0-alpha | 2025-11-03 | Initial MVP with all core modules and GUI |

---

## Commit History

### Recent Commits (1.0.0-beta)

- `3b00e66` - Fix bare except blocks across codebase
- `21ac58e` - Add comprehensive type hints to priority modules
- `6ca55bf` - Refactor workflow_helpers.py into focused modules
- `8a82847` - Fix build: Add resources directory and fix Unicode in build script
- `2930185` - Fix Excel format_cells test to use proper range syntax
- `62a2d35` - Fix Excel fill color test assertion
- `1f27cd5` - Fix Python < 3.12 f-string backslash compatibility
- `38dd3ae` - Fix CI: Windows-only testing and Python 3.10+ compatibility
- `9a33253` - Fix CI pipeline and test compatibility issues
- `4770b77` - Add GitHub Actions CI/CD pipeline
- `4bdf0e4` - Add comprehensive pytest test suite
- `9339da8` - Organize test files into proper test directory structure
- `d692ce7` - Add example automation scripts
- `b5ad089` - Update Word automation document handler
- `0466082` - Add Asana integration module with browser automation
- `cadc0ff` - Enhance OneNote integration with COM client and content formatting
- `7284ba7` - Add AI workflow generator with Claude API integration
- `b800a71` - Implement Task Scheduler with APScheduler integration
- `140d743` - Implement Script Manager, Execution Engine, and Script Dialog

### Initial Commits (1.0.0-alpha)

- `ba27157` - Fix TypeError in get_recent_logs when handling Path objects
- `c191327` - Add comprehensive project status document
- `9685100` - Initial commit: Automation Hub MVP

---

## Migration Notes

### From 1.0.0-alpha to 1.0.0-beta

**Breaking Changes:** None

**New Dependencies:**
- `pytest` >= 7.0.0
- `pytest-cov` >= 4.0.0
- `pytest-qt` >= 4.2.0
- `anthropic` >= 0.18.0 (optional, for AI features)
- `apscheduler` >= 3.10.0

**Configuration Changes:**
- New `ai` section in config for Claude API settings
- New `onenote` section for default notebook/section

**File Structure Changes:**
- `src/utils/workflow_helpers.py` split into 5 focused modules
- New `tests/` directory structure with unit and integration tests
- New `.github/workflows/` for CI/CD configuration

---

## Contributors

- Development Team
- Claude AI (Code assistance)

---

## Links

- [GitHub Repository](https://github.com/criswell82/automation-hub)
- [Documentation](docs/)
- [Issue Tracker](https://github.com/criswell82/automation-hub/issues)
