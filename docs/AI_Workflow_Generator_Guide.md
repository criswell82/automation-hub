# AI Workflow Generator - User Guide

## Overview

The AI Workflow Generator allows you to create custom automation workflows by simply describing what you want to automate in plain English. The system uses AI to generate Python scripts that integrate with Automation Hub's modules.

## Features

✨ **Natural Language Creation**: Describe your workflow in plain English
🤖 **AI-Powered Generation**: Automatically generates Python code
📝 **Customizable**: Review and modify generated code before saving
🔄 **Instant Integration**: Generated workflows appear immediately in your script library
📚 **Template Library**: Pre-built templates for common tasks

## Getting Started

### Prerequisites

For AI-powered generation, you need:
- An Anthropic API key (Claude AI)
- The `anthropic` Python package installed

### Setting Up API Key

**Option 1: Environment Variable (Recommended)**
```bash
# Windows
set ANTHROPIC_API_KEY=your_api_key_here

# Or add to system environment variables permanently
```

**Option 2: Configuration File**
Add to your config file at `%APPDATA%/AutomationHub/config/config.json`:
```json
{
  "ai": {
    "anthropic_api_key": "your_api_key_here"
  }
}
```

**Note**: Without an API key, the system will fall back to template-based generation.

## Using the Workflow Generator

### Step 1: Open the Generator

1. Launch Automation Hub
2. Click `File` → `Create New Workflow` (or press `Ctrl+N`)

### Step 2: Describe Your Workflow

In the description box, explain what you want to automate in plain English. Be specific!

**Good Examples:**

```
Create an Excel report from CSV files in C:/Data folder,
add a summary chart, and email it to manager@company.com
every Monday morning.
```

```
Process all unread emails in my inbox, extract PDF attachments,
and save them to SharePoint organized by date.
```

```
Organize files in my Downloads folder by file type,
archive files older than 90 days, and sync to SharePoint.
```

**Tips for Better Results:**
- Be specific about file paths and locations
- Mention exact email addresses or SharePoint URLs if applicable
- Specify scheduling requirements if needed
- Include data processing steps

### Step 3: Configure Options

- **Category**: Select the workflow category or choose "Auto-detect"
- **Use template examples**: Check this for better results (includes examples in generation)

### Step 4: Generate

1. Click **Generate Workflow**
2. Wait while the AI analyzes and generates code (usually 5-15 seconds)
3. Review the generated code in the preview panel

### Step 5: Save and Use

- **Save Workflow**: Saves to `user_scripts/custom/` directory
- **Save & Run**: Saves and immediately opens run dialog
- **Close**: Cancel without saving

After saving, the workflow will automatically appear in your script library!

## Workflow Structure

Generated workflows follow this structure:

```python
"""
WORKFLOW_META:
  name: Your Workflow Name
  description: What it does
  category: Reports/Email/Files/etc
  version: 1.0.0
  author: Automation Hub
  parameters:
    param_name:
      type: string/file/choice/boolean
      description: Parameter help text
      required: true/false
      default: "default value"
"""

import sys
from pathlib import Path

# Setup Python path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)

class YourWorkflow(BaseModule):
    def configure(self, **kwargs):
        # Setup parameters
        pass

    def validate(self):
        # Validate configuration
        return True

    def execute(self):
        # Main workflow logic
        try:
            logger.info("Starting workflow...")
            # Your automation code here
            return {"status": "success"}
        except Exception as e:
            self.handle_error(e, "Workflow failed")
            raise

def run(**kwargs):
    workflow = YourWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
```

## Pre-Built Templates

The system includes ready-to-use templates for common scenarios:

### 1. Weekly Report Generator
**Location**: `user_scripts/templates/reports/weekly_report_generator.py`

**What it does:**
- Collects data from CSV/Excel files in a folder
- Generates formatted Excel report with charts
- Emails report to recipients
- Perfect for weekly/monthly reporting tasks

**Parameters:**
- `data_folder`: Folder containing source data files
- `output_file`: Where to save the Excel report
- `email_to`: Recipients (comma-separated)
- `email_subject`: Email subject line

### 2. Email Attachment Processor
**Location**: `user_scripts/templates/email/email_attachment_processor.py`

**What it does:**
- Processes emails from Outlook
- Extracts attachments (filtered by file type)
- Organizes by date
- Great for handling invoice PDFs, reports, etc.

**Parameters:**
- `email_folder`: Outlook folder to process (default: Inbox)
- `subject_filter`: Filter by subject keyword
- `output_folder`: Where to save attachments
- `file_types`: Extensions to extract (.pdf,.xlsx, etc.)
- `unread_only`: Only process unread emails
- `organize_by_date`: Organize into date folders

### 3. File Organizer and SharePoint Sync
**Location**: `user_scripts/templates/files/file_organizer_and_sync.py`

**What it does:**
- Organizes files by type or date
- Archives old files
- Syncs to SharePoint
- Perfect for keeping downloads/folders tidy

**Parameters:**
- `source_folder`: Folder to organize
- `destination_folder`: Where to put organized files
- `organize_method`: by_type, by_date, or both
- `archive_old_files`: Archive files older than 90 days
- `sync_to_sharepoint`: Upload to SharePoint
- `sharepoint_site`: SharePoint site URL
- `sharepoint_folder`: SharePoint folder path

## Helper Functions Library

The `workflow_helpers.py` module provides ready-to-use components:

### Report Generation

```python
from src.utils.workflow_helpers import ReportBuilder

report = ReportBuilder()
report.add_title_sheet("My Report", "Subtitle")
report.add_data_sheet("Data", data_rows, headers)
report.add_summary_with_chart("Summary", summary_dict, chart_type='bar')
report.save("output.xlsx")
```

### Email Automation

```python
from src.utils.workflow_helpers import EmailHelper

email = EmailHelper()
email.send_report(
    to=['recipient@example.com'],
    subject='Report',
    body='Please find attached...',
    attachments=['report.xlsx']
)

emails = email.process_inbox(folder='Inbox', unread_only=True)
files = email.extract_attachments(emails, 'C:/Attachments')
```

### File Management

```python
from src.utils.workflow_helpers import FileOrganizer

organizer = FileOrganizer()
organizer.organize_by_type('C:/Downloads', 'C:/Organized')
organizer.organize_by_date('C:/Files', 'C:/ByDate', date_format="%Y-%m")
organizer.archive_old_files('C:/Files', 'C:/Archive', days_old=90)
organizer.batch_rename('C:/Files', 'old_name', 'new_name')
```

### SharePoint Integration

```python
from src.utils.workflow_helpers import SharePointHelper

sp = SharePointHelper('https://company.sharepoint.com/sites/yoursite')
sp.sync_folder_to_sharepoint('C:/Local', '/Shared Documents/Folder')
sp.download_recent_files('/Shared Documents', 'C:/Local', days=7)
```

## Customizing Generated Workflows

You can edit generated workflows manually:

1. Open the file in `user_scripts/custom/` with any text editor
2. Modify the code as needed
3. Save the file
4. Click **Scripts** → **Refresh Scripts** in Automation Hub

Common customizations:
- Add additional parameters
- Change data processing logic
- Add error handling
- Integrate additional modules
- Modify output formats

## Scheduling Workflows

Once created, you can schedule your workflows:

1. Select the workflow in the script library
2. Click **Scripts** → **Schedule Script** (or `Ctrl+S`)
3. Configure schedule:
   - **One-time**: Run at a specific date/time
   - **Interval**: Run every N minutes/hours/days
   - **Cron**: Advanced scheduling (e.g., "0 9 * * 1" for Mondays at 9 AM)
4. Click **Add Schedule**

Your workflow will run automatically according to the schedule!

## Troubleshooting

### Generation Fails

**Problem**: "Failed to generate workflow"

**Solutions:**
- Check your API key is valid
- Ensure you have internet connection
- Try again with a more specific description
- Fall back to using templates directly

### Workflow Doesn't Appear

**Problem**: Saved workflow doesn't show in library

**Solutions:**
- Click **Scripts** → **Refresh Scripts** (or `F5`)
- Check the file was saved to `user_scripts/custom/`
- Verify the WORKFLOW_META section is properly formatted

### Execution Errors

**Problem**: Workflow fails when running

**Solutions:**
- Check file paths exist and are correct
- Verify you have permissions for folders
- Review logs in the Output tab
- Check parameter values are correct

### Missing Dependencies

**Problem**: Import errors when running workflow

**Solutions:**
- Ensure all required modules are installed
- For AI generation: `pip install anthropic`
- For SharePoint: Office365 credentials configured
- Check `requirements.txt` for all dependencies

## Best Practices

### 1. Start with Templates

Begin with pre-built templates and customize them for your needs. This is faster than generating from scratch.

### 2. Test Before Scheduling

Always test workflows manually before scheduling them:
1. Run with sample data
2. Verify outputs are correct
3. Check email sending works
4. Then schedule for production

### 3. Use Descriptive Names

Give workflows clear, descriptive names:
- ✅ "Weekly Sales Report to Manager"
- ✅ "Process Invoice PDFs from Inbox"
- ❌ "Workflow1"
- ❌ "Test"

### 4. Document Parameters

Add clear descriptions to parameters so you remember what they do later.

### 5. Handle Errors Gracefully

Generated workflows include error handling. Customize error messages for your use case.

### 6. Keep Workflows Focused

Create separate workflows for different tasks rather than one giant workflow. This makes debugging and maintenance easier.

## Examples by Use Case

### Daily Sales Report

**Description:**
```
Generate a daily sales report from C:/Sales/Daily/*.csv files,
create an Excel report with charts, and email it to sales@company.com
every day at 8 AM.
```

**Result**: Workflow that processes daily sales data, creates formatted report, and emails it automatically.

### Invoice Processing

**Description:**
```
Monitor my Outlook inbox for emails with "Invoice" in the subject,
extract PDF attachments, save them to C:/Invoices organized by month,
and mark emails as read.
```

**Result**: Workflow that automatically processes invoice emails and organizes attachments.

### Download Folder Cleanup

**Description:**
```
Organize my Downloads folder by file type (Documents, Images, etc.),
archive files older than 3 months to C:/Archive, and sync
important documents to SharePoint.
```

**Result**: Workflow that keeps Downloads tidy and backs up important files.

## Getting Help

- **Documentation**: Check `docs/` folder for detailed guides
- **Examples**: See `user_scripts/templates/` for working examples
- **Logs**: Review the Output tab for error messages
- **Code**: Generated workflows include comments explaining each step

## Advanced Topics

### Using Multiple Modules

Combine different automation modules in one workflow:

```python
# Excel + Email
from src.modules.excel_automation import WorkbookHandler
from src.modules.outlook_automation import EmailHandler

# Generate report
wb = WorkbookHandler()
# ... create report ...
wb.save('report.xlsx')

# Email it
email = EmailHandler()
email.send_email(..., attachments=['report.xlsx'])
```

### Custom Data Processing

Add your own data transformation logic:

```python
def process_data(data):
    # Your custom processing
    filtered = [row for row in data if row['amount'] > 100]
    totals = sum(row['amount'] for row in filtered)
    return {'filtered': filtered, 'total': totals}
```

### Integration with External APIs

Import additional libraries as needed:

```python
import requests

def fetch_external_data():
    response = requests.get('https://api.example.com/data')
    return response.json()
```

## Next Steps

1. **Try the Generator**: Create your first workflow using the AI generator
2. **Explore Templates**: Check out the pre-built templates
3. **Customize**: Modify templates to match your exact needs
4. **Schedule**: Set up automatic execution for recurring tasks
5. **Share**: Export workflows to share with team members (coming soon!)

---

**Questions?** Check the main Automation Hub documentation or review example workflows in `user_scripts/templates/`.
