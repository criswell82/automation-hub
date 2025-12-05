# Asana Integration Guide

**Comprehensive guide to integrating Automation Hub with Asana without API access**

Version: 1.0.0
Last Updated: November 10, 2024

---

## Table of Contents

1. [Overview](#overview)
2. [Why No API Access?](#why-no-api-access)
3. [Integration Methods](#integration-methods)
4. [Quick Start](#quick-start)
5. [Module Reference](#module-reference)
6. [Workflow Templates](#workflow-templates)
7. [Use Cases & Examples](#use-cases--examples)
8. [Troubleshooting](#troubleshooting)
9. [Best Practices](#best-practices)

---

## Overview

The Asana integration allows you to automate task creation, bulk imports, roadmap management, and more **without requiring Asana API access**. This is perfect for corporate environments with restricted API access or teams that want simple, reliable automation.

### Key Features

✅ **No API Keys Required** - Uses email-to-task and CSV import
✅ **Bulk Operations** - Create 100s of tasks from Excel files
✅ **Multiple Methods** - Email, CSV, or browser automation
✅ **Corporate-Friendly** - Works within IT restrictions
✅ **Pre-built Templates** - Ready-to-use workflows
✅ **AI-Generated Workflows** - Describe what you want in plain English

---

## Why No API Access?

Many organizations restrict API access due to:
- Security policies (no external API keys allowed)
- Asana plan limitations (API requires paid plans)
- IT approval processes (too slow for quick automation needs)
- Preference for simpler methods

Our integration works around these constraints by using:
1. **Email-to-Task**: Every Asana project has a unique email address
2. **CSV Import/Export**: Asana's built-in bulk import feature
3. **Browser Automation**: RPA for complex operations

---

## Integration Methods

### Method 1: Email-to-Task (Recommended ⭐)

**How it works:**
- Send email to project's unique address
- Email subject → Task name
- Email body → Task description
- Special syntax for assignee, due date, priority

**Pros:**
- ✅ Very reliable (email rarely fails)
- ✅ No API keys needed
- ✅ Works offline from Asana
- ✅ Can batch 100s of tasks
- ✅ Corporate firewall-friendly

**Cons:**
- ❌ Cannot read existing tasks
- ❌ One-way sync only
- ❌ Need project email addresses

**Best for:**
- Daily/weekly task creation
- Sprint planning automation
- Bulk imports from Excel
- Automated workflows

---

### Method 2: CSV Import/Export

**How it works:**
- Generate Asana-compatible CSV file
- Manually import in Asana UI (or automate with RPA)
- Can handle 1000s of tasks at once

**Pros:**
- ✅ Massive scale (1000s of tasks)
- ✅ Great for project setup
- ✅ Review before import
- ✅ Column mapping flexibility

**Cons:**
- ❌ Manual import step required
- ❌ No real-time sync

**Best for:**
- Quarterly roadmap imports
- Project initialization
- Large migrations
- One-time bulk operations

---

### Method 3: Browser Automation (Advanced)

**How it works:**
- Automates Asana web interface using RPA
- Can read and write tasks
- Full UI control via mouse/keyboard

**Pros:**
- ✅ Can read existing tasks
- ✅ Two-way sync possible
- ✅ Access all Asana features

**Cons:**
- ❌ Slower than email/CSV
- ❌ UI changes can break automation
- ❌ Requires browser visibility

**Best for:**
- Complex workflows
- Bidirectional sync
- When other methods insufficient

---

## Quick Start

### Step 1: Find Your Project Email

1. Open your Asana project in browser
2. Click project dropdown (▼) next to project name
3. Go to **Settings**
4. Scroll to **"Email address for this project"**
5. Copy the email (e.g., `my-project-1234567890@mail.asana.com`)

### Step 2: Create Your First Task

```python
from src.utils.workflow_helpers import AsanaHelper

# Initialize helper
helper = AsanaHelper(project_email='your-project@mail.asana.com')

# Create a task
result = helper.create_task_via_email(
    name='Review Q4 Budget',
    description='Review and approve the Q4 budget proposal',
    assignee='john.doe',
    due_date='2024-12-15',
    priority='High'
)

print(result)
```

### Step 3: Bulk Import from Excel

**Create Excel file** (`tasks.xlsx`):

| name | description | assignee | due_date | priority | tags |
|------|-------------|----------|----------|----------|------|
| Task 1 | Description | john.doe | 2024-12-15 | High | feature |
| Task 2 | Another task | jane.smith | 2024-12-20 | Medium | bug |

**Run the import:**

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper(project_email='your-project@mail.asana.com')
result = helper.bulk_create_from_excel('C:/Data/tasks.xlsx')

print(f"Created {result['data']['created_count']} tasks!")
```

---

## Module Reference

### AsanaEmailModule

Create tasks via email-to-task integration.

```python
from src.modules.asana import AsanaEmailModule

module = AsanaEmailModule()
result = module.run(
    project_email='project@mail.asana.com',
    tasks=[
        {
            'name': 'Task name',
            'description': 'Task description',
            'assignee': 'username',
            'due_date': '2024-12-15',
            'priority': 'High',
            'tags': 'feature,sprint-42'
        }
    ]
)
```

**Task Dictionary Fields:**
- `name` (required): Task name
- `description`: Task notes/description
- `assignee`: Username (without @)
- `due_date`: YYYY-MM-DD or "next Friday"
- `priority`: High/Medium/Low (adds !! or ! to name)
- `tags`: Comma-separated tags
- `attachments`: List of file paths

---

### AsanaCSVHandler

Generate and parse Asana CSV files.

```python
from src.modules.asana import AsanaCSVHandler

# Generate CSV for import
handler = AsanaCSVHandler()
result = handler.run(
    operation='generate',
    tasks=[...],  # List of task dictionaries
    output_file='C:/Output/asana_import.csv'
)

# Parse Asana export
result = handler.run(
    operation='parse',
    input_file='C:/Downloads/asana_export.csv',
    output_file='C:/Output/tasks.json'
)

# Convert custom CSV to Asana format
result = handler.run(
    operation='convert',
    input_file='C:/Data/my_tasks.csv',
    output_file='C:/Output/asana_import.csv',
    mapping={
        'Task Name': 'title',
        'Assignee': 'owner',
        'Due Date': 'deadline'
    }
)
```

---

### AsanaBrowserModule

Automate Asana UI via browser RPA.

```python
from src.modules.asana import AsanaBrowserModule

module = AsanaBrowserModule()

# Create tasks via browser automation
result = module.run(
    asana_url='https://app.asana.com/0/1234567890/list',
    operation='create_tasks',
    tasks=[
        {'name': 'Task 1'},
        {'name': 'Task 2'}
    ],
    wait_time=3  # Seconds to wait for page loads
)

# Read tasks (template - requires customization)
result = module.run(
    asana_url='https://app.asana.com/0/1234567890/list',
    operation='read_tasks'
)
```

---

### AsanaHelper

High-level helper class for common operations.

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper(
    project_email='project@mail.asana.com',
    project_url='https://app.asana.com/0/1234567890/list'
)

# Create single task
helper.create_task_via_email(
    name='Task name',
    description='Description',
    assignee='john.doe',
    due_date='2024-12-15'
)

# Bulk create from Excel
helper.bulk_create_from_excel('tasks.xlsx', method='email')

# Generate CSV for manual import
helper.generate_asana_csv(tasks, 'output.csv')

# Parse Asana export
tasks = helper.parse_asana_export('asana_export.csv')

# Sync Excel tracker to Asana
helper.sync_excel_tracker_to_asana('tracker.xlsx')

# Create weekly sprint tasks
helper.create_weekly_sprint_tasks(
    sprint_number=42,
    start_date='2024-12-09',
    task_templates=[...]
)
```

---

## Workflow Templates

Pre-built workflows in `user_scripts/templates/asana/`:

### 1. Excel to Asana (`excel_to_asana.py`)

Bulk import tasks from Excel file.

```python
from user_scripts.templates.asana.excel_to_asana import ExcelToAsanaWorkflow

workflow = ExcelToAsanaWorkflow()
result = workflow.run(
    project_email='project@mail.asana.com',
    excel_file='C:/Data/tasks.xlsx'
)
```

---

### 2. Weekly Sprint Tasks (`weekly_sprint_tasks.py`)

Auto-create weekly sprint tasks.

```python
from user_scripts.templates.asana.weekly_sprint_tasks import WeeklySprintTasksWorkflow

workflow = WeeklySprintTasksWorkflow()
result = workflow.run(
    project_email='sprint@mail.asana.com',
    sprint_number=42,
    start_date='2024-12-09'
)
```

**Creates:**
- Sprint Planning (Monday)
- Daily Standups (Mon-Fri)
- Sprint Review (Friday)
- Sprint Retrospective (Friday)
- Update Sprint Metrics (Friday)

---

### 3. Generate Asana CSV (`generate_asana_csv.py`)

Convert Excel to Asana CSV format.

```python
from user_scripts.templates.asana.generate_asana_csv import GenerateAsanaCSVWorkflow

workflow = GenerateAsanaCSVWorkflow()
result = workflow.run(
    input_file='C:/Data/tasks.xlsx',
    output_csv='C:/Output/asana_import.csv'
)
```

---

### 4. Quarterly Roadmap (`quarterly_roadmap_import.py`)

Import quarterly roadmap from Excel.

```python
from user_scripts.templates.asana.quarterly_roadmap_import import QuarterlyRoadmapWorkflow

workflow = QuarterlyRoadmapWorkflow()
result = workflow.run(
    project_email='q1@mail.asana.com',
    roadmap_file='C:/Planning/Q1_Roadmap.xlsx',
    quarter='Q1 2025'
)
```

---

## Use Cases & Examples

### Use Case 1: Daily Sprint Planning

**Scenario:** Automatically create daily standup tasks every morning.

```python
from src.utils.workflow_helpers import AsanaHelper
from datetime import datetime

helper = AsanaHelper(project_email='sprint@mail.asana.com')

today = datetime.now().strftime('%Y-%m-%d')
day_name = datetime.now().strftime('%A')

helper.create_task_via_email(
    name=f'Daily Standup - {day_name}',
    description='What did you do? What will you do? Blockers?',
    assignee='team',
    due_date=today,
    tags='standup'
)
```

**Schedule this to run every weekday at 9 AM using Automation Hub's scheduler!**

---

### Use Case 2: Convert Project Tracker to Asana

**Scenario:** You have an Excel tracker, want to migrate to Asana.

```python
from src.utils.workflow_helpers import create_asana_roadmap_from_excel

result = create_asana_roadmap_from_excel(
    excel_path='C:/Tracking/project_tracker.xlsx',
    project_email='project@mail.asana.com'
)

print(f"Imported {result['data']['created_count']} tasks!")
```

---

### Use Case 3: Weekly Roadmap Sync

**Scenario:** Update Asana roadmap every Monday from planning Excel file.

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper(project_email='roadmap@mail.asana.com')

# Generate CSV for review
helper.convert_excel_to_asana_csv(
    excel_path='C:/Planning/Weekly_Roadmap.xlsx',
    output_path='C:/Output/roadmap_import.csv'
)

print("Review the CSV, then import manually in Asana")
```

---

### Use Case 4: Task Status Reporting

**Scenario:** Export Asana tasks, analyze in Excel.

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper()

# Parse Asana export
tasks = helper.parse_asana_export('C:/Downloads/asana_export.csv')

# Analyze
total = len(tasks)
completed = len([t for t in tasks if t.get('completed_at')])
in_progress = total - completed

print(f"Total: {total}")
print(f"Completed: {completed} ({completed/total*100:.1f}%)")
print(f"In Progress: {in_progress}")
```

---

## Troubleshooting

### Tasks Not Appearing in Asana

**Problem:** Sent emails but tasks don't appear.

**Solutions:**
1. **Verify project email**: Check Settings for correct address
2. **Check Outlook Sent folder**: Ensure emails actually sent
3. **Check Asana spam**: New emails might be filtered
4. **Wait 1-2 minutes**: Sometimes there's a delay
5. **Check email format**: Subject must not be empty

---

### Excel File Not Found

**Problem:** `FileNotFoundError: Excel file not found`

**Solutions:**
1. Use absolute paths: `C:/Data/tasks.xlsx` not `tasks.xlsx`
2. Check file exists at that location
3. Close Excel if file is open
4. Check file extension (.xlsx vs .xls)

---

### pywin32 Not Available

**Problem:** `pywin32 is required but not installed`

**Solutions:**
```bash
pip install pywin32
```

Or use CSV method instead:
```python
helper.bulk_create_from_excel('file.xlsx', method='csv')
```

---

### Assignee Not Recognized

**Problem:** Tasks created but assignee not set.

**Solutions:**
1. Use exact Asana username (check in profile)
2. Don't include @ symbol: `john.doe` not `@john.doe`
3. Case-sensitive: match Asana exactly
4. For guest users, use their email address

---

### Too Many Emails

**Problem:** Outlook sending 100+ emails feels spammy.

**Solutions:**
1. Use CSV method for large bulk operations:
   ```python
   helper.generate_asana_csv(tasks, 'import.csv')
   ```
2. Batch emails by sending fewer tasks per run
3. Use delay between emails:
   ```python
   import time
   for task in tasks:
       helper.create_task_via_email(**task)
       time.sleep(2)  # 2 second delay
   ```

---

## Best Practices

### 1. Start Small

- Test with 5-10 tasks first
- Verify they appear correctly in Asana
- Then scale to larger operations

### 2. Use Descriptive Names

```python
# ✅ Good
name='Sprint 42: Code Review - User Auth Module'

# ❌ Bad
name='Review'
```

### 3. Add Context in Description

```python
description='''
Review the user authentication module PR.

Checklist:
- [ ] Security best practices
- [ ] Error handling
- [ ] Test coverage
- [ ] Documentation
'''
```

### 4. Use Tags for Organization

```python
tags='sprint-42,backend,code-review'
```

### 5. Schedule Recurring Tasks

Use Automation Hub's scheduler for:
- Daily standups (every weekday 9 AM)
- Weekly sprint planning (every Monday)
- Monthly roadmap sync (first of month)

### 6. Excel File Best Practices

**Column order doesn't matter**, but use these exact names:
- `name` (required)
- `description`
- `assignee`
- `due_date` (YYYY-MM-DD format)
- `priority` (High/Medium/Low)
- `tags` (comma-separated)
- `section`

**Example:**
```
name,assignee,due_date,priority,tags
Review Q4 Budget,john.doe,2024-12-15,High,finance
Update Documentation,jane.smith,2024-12-20,Medium,docs
```

### 7. Project Email Organization

Keep a config file with all project emails:

```python
# asana_projects.py
PROJECTS = {
    'sprint': 'sprint-team@mail.asana.com',
    'roadmap': 'product-roadmap@mail.asana.com',
    'bugs': 'bug-tracking@mail.asana.com',
    'q1': 'q1-2025@mail.asana.com'
}

# Usage
from asana_projects import PROJECTS
helper = AsanaHelper(project_email=PROJECTS['sprint'])
```

---

## Advanced Topics

### Custom Task Templates

Create reusable task templates:

```python
TASK_TEMPLATES = {
    'code_review': {
        'name': 'Code Review: {pr_number}',
        'description': 'Review PR #{pr_number}\n\nChecklist:\n- [ ] Code quality\n- [ ] Tests\n- [ ] Docs',
        'priority': 'High',
        'tags': 'code-review'
    },
    'bug_fix': {
        'name': 'Fix Bug: {bug_description}',
        'description': 'Bug report: {bug_description}\n\nSteps to reproduce:\n{steps}',
        'priority': 'High',
        'tags': 'bug'
    }
}

# Use template
template = TASK_TEMPLATES['code_review'].copy()
template['name'] = template['name'].format(pr_number=123)
helper.create_task_via_email(**template)
```

### AI-Generated Workflows

Use the Workflow Generator to create custom Asana workflows:

1. Open Automation Hub
2. Click "Workflow Generator"
3. Describe in plain English:
   ```
   Create a workflow that reads tasks from Excel,
   adds a sprint tag, and creates them in Asana
   with high priority for anything due this week.
   ```
4. AI generates the code automatically!

---

## API Reference Summary

### AsanaHelper Methods

| Method | Purpose | Returns |
|--------|---------|---------|
| `create_task_via_email()` | Create single task | Result dict |
| `bulk_create_from_list()` | Create multiple tasks | Result dict |
| `bulk_create_from_excel()` | Import from Excel | Result dict |
| `generate_asana_csv()` | Generate CSV file | Result dict |
| `parse_asana_export()` | Parse Asana CSV | List of tasks |
| `convert_excel_to_asana_csv()` | Convert formats | Result dict |
| `sync_excel_tracker_to_asana()` | Sync tracker | Result dict |
| `create_weekly_sprint_tasks()` | Sprint automation | Result dict |

### Common Workflow Functions

| Function | Purpose |
|----------|---------|
| `create_asana_roadmap_from_excel()` | Import roadmap |
| `sync_asana_status_to_excel()` | Bidirectional sync |

---

## Support & Resources

### Documentation
- **Main docs**: `docs/user_guide.md`
- **Template README**: `user_scripts/templates/asana/README.md`
- **Developer guide**: `docs/developer_guide.md`

### Examples
- Browse templates in `user_scripts/templates/asana/`
- Check example scripts in module docstrings
- Use AI Workflow Generator for custom scripts

### Getting Help
1. Review this guide and template README
2. Check troubleshooting section
3. Review error messages carefully
4. Test with small datasets first

---

## Version History

### 1.0.0 (2024-11-10)
- Initial release
- Email-to-task integration
- CSV import/export
- Browser automation (template)
- AsanaHelper convenience class
- 4 pre-built workflow templates
- AI workflow generator support
- Comprehensive documentation

---

## What's Next?

Explore these advanced features:
1. Schedule recurring Asana workflows
2. Build custom templates for your team
3. Integrate with other modules (Excel → Asana → Email reports)
4. Use AI to generate custom workflows
5. Create bidirectional sync with Excel trackers

**Happy Automating! 🚀**
