# Automation Hub - User Guide

## Table of Contents

1. [Getting Started](#getting-started)
2. [The Dashboard](#the-dashboard)
3. [Running Scripts](#running-scripts)
4. [Scheduling Tasks](#scheduling-tasks)
5. [Configuration](#configuration)
6. [Troubleshooting](#troubleshooting)

## Getting Started

### First Launch

When you first launch Automation Hub, you'll be greeted with the main dashboard. The application will create necessary configuration files in your user directory.

### Initial Setup

1. **Review Settings**: Click Settings to configure basic preferences
2. **Test Credentials**: Set up any required credentials for SharePoint, Graph API, etc.
3. **Browse Scripts**: Explore available automation scripts in the script library

## The Dashboard

The main dashboard consists of several key areas:

### Script Library
Browse all available automation scripts organized by module:
- Desktop RPA
- Excel Automation
- Outlook Automation
- SharePoint Integration
- Word Automation
- OneNote Integration

### Quick Actions
- **Recent Scripts**: Access your most recently used scripts
- **Favorites**: Pin frequently used scripts for quick access
- **Scheduled Tasks**: View and manage scheduled executions

### Output Viewer
Real-time display of script execution output and logs with:
- Filtering by log level (INFO, WARNING, ERROR)
- Search functionality
- Export to file

### Status Monitor
Track the status of running scripts:
- Currently executing
- Queued for execution
- Recently completed
- Failed executions

## Running Scripts

### Manual Execution

1. Navigate to the script in the Script Library
2. Click the script name to view details
3. Configure any required parameters
4. Click **Run Now** to execute immediately

### Script Parameters

Many scripts require configuration parameters such as:
- File paths
- Date ranges
- Email addresses
- SharePoint URLs

The parameter configuration dialog will guide you through required inputs.

## Scheduling Tasks

### Creating a Schedule

1. Select a script from the library
2. Click **Schedule** instead of **Run Now**
3. Choose schedule type:
   - **One-time**: Execute at a specific date/time
   - **Daily**: Repeat every day at a specific time
   - **Weekly**: Repeat on specific days of the week
   - **Monthly**: Repeat on specific days of the month

### Managing Schedules

View all scheduled tasks in the **Scheduled Tasks** panel:
- Edit schedule parameters
- Temporarily disable schedules
- Delete schedules
- View execution history

## Configuration

### Application Settings

Access via **File > Settings**:

- **General**: Application preferences
- **Logging**: Log level and retention settings
- **Security**: Credential management
- **Updates**: Auto-update configuration

### Module Configuration

Each automation module has specific configuration options:

#### Desktop RPA
- Default delay between actions
- Window detection timeout
- Screenshot save location

#### Excel Automation
- Default Excel template location
- Chart style preferences
- Number formatting

#### Outlook Automation
- Email signature
- Default folder paths
- Task creation settings

#### SharePoint
- Default site URL
- Authentication method
- Download location

## Troubleshooting

### Common Issues

**Script fails to execute**
- Check the output viewer for error messages
- Verify all required parameters are provided
- Ensure necessary permissions are available

**Scheduled task didn't run**
- Verify the computer was powered on at scheduled time
- Check Windows Task Scheduler integration
- Review execution logs

**Authentication errors**
- Re-enter credentials in Settings
- Verify network connectivity
- Check with IT for access permissions

### Getting Help

- View detailed error messages in the Output Viewer
- Check script-specific documentation
- Contact development team for support

## Best Practices

1. **Test First**: Run scripts manually before scheduling
2. **Monitor Logs**: Review logs regularly for issues
3. **Backup Configuration**: Export settings before major changes
4. **Update Regularly**: Keep Automation Hub updated
5. **Document Custom Scripts**: Add clear descriptions to custom automations

## Keyboard Shortcuts

- `Ctrl+R`: Run selected script
- `Ctrl+S`: Schedule selected script
- `Ctrl+F`: Search scripts
- `Ctrl+L`: Clear output viewer
- `F5`: Refresh script library
- `Ctrl+,`: Open settings

## Safety Features

Automation Hub includes several safety features:
- Confirmation prompts for destructive actions
- Dry-run mode for testing
- Automatic backup before overwrites
- Execution history and rollback capabilities

---

*For technical documentation and custom script development, see the [Developer Guide](developer_guide.md).*
