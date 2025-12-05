# Asana Integration Templates

These templates provide ready-to-use workflows for integrating with Asana **without requiring API access**.

## Available Templates

### 1. Excel to Asana Bulk Import
**File:** `excel_to_asana.py`

Bulk create Asana tasks from an Excel file using email-to-task integration.

**Excel Format:**
| name | description | assignee | due_date | priority | tags |
|------|-------------|----------|----------|----------|------|
| Task 1 | Description here | john.doe | 2024-12-15 | High | feature |
| Task 2 | Another task | jane.smith | 2024-12-20 | Medium | bug-fix |

**Usage:**
```python
from excel_to_asana import ExcelToAsanaWorkflow

workflow = ExcelToAsanaWorkflow()
result = workflow.run(
    project_email='your-project@mail.asana.com',
    excel_file='C:/Data/tasks.xlsx'
)
```

---

### 2. Weekly Sprint Tasks
**File:** `weekly_sprint_tasks.py`

Automatically create weekly sprint tasks from predefined templates.

**Usage:**
```python
from weekly_sprint_tasks import WeeklySprintTasksWorkflow

workflow = WeeklySprintTasksWorkflow()
result = workflow.run(
    project_email='sprint@mail.asana.com',
    sprint_number=42,
    start_date='2024-12-09'  # Sprint start date
)
```

**Creates:**
- Sprint Planning (Monday)
- Daily Standups (Mon-Fri)
- Sprint Review (Friday)
- Sprint Retrospective (Friday)
- Update Sprint Metrics (Friday)

---

### 3. Generate Asana Import CSV
**File:** `generate_asana_csv.py`

Convert Excel or custom CSV to Asana-compatible CSV format for manual import.

**Usage:**
```python
from generate_asana_csv import GenerateAsanaCSVWorkflow

workflow = GenerateAsanaCSVWorkflow()
result = workflow.run(
    input_file='C:/Data/project_tasks.xlsx',
    output_csv='C:/Output/asana_import.csv',
    column_mapping='{"Task Name": "title", "Assignee": "owner"}'
)
```

Then import the generated CSV manually in Asana:
1. Open project → Click dropdown → Import → CSV
2. Select the generated file
3. Map columns and import

---

### 4. Quarterly Roadmap Import
**File:** `quarterly_roadmap_import.py`

Import quarterly roadmap from Excel planning file into Asana.

**Usage:**
```python
from quarterly_roadmap_import import QuarterlyRoadmapWorkflow

workflow = QuarterlyRoadmapWorkflow()
result = workflow.run(
    project_email='q1-roadmap@mail.asana.com',
    roadmap_file='C:/Planning/Q1_Roadmap.xlsx',
    quarter='Q1 2025'
)
```

---

## Getting Started

### Step 1: Find Your Project Email

Every Asana project has a unique email address:

1. Open your Asana project
2. Click the dropdown next to project name
3. Go to **Settings**
4. Scroll to **Email address for this project**
5. Copy the email (e.g., `your-project-1234567890@mail.asana.com`)

### Step 2: Prepare Your Excel File

Create an Excel file with these columns (all optional except `name`):

- **name** (required): Task name
- **description**: Task description/notes
- **assignee**: Asana username (without @)
- **due_date**: Due date in YYYY-MM-DD format
- **priority**: High, Medium, or Low
- **tags**: Comma-separated tags
- **section**: Project section name

### Step 3: Run a Workflow

**Option A: Using the GUI**
1. Open Automation Hub
2. Go to Script Library
3. Find the Asana template
4. Configure parameters
5. Click Run

**Option B: Using Python**
```python
import sys
sys.path.append('path/to/automation_hub')

from user_scripts.templates.asana.excel_to_asana import ExcelToAsanaWorkflow

workflow = ExcelToAsanaWorkflow()
result = workflow.run(
    project_email='your-project@mail.asana.com',
    excel_file='C:/Data/tasks.xlsx'
)

print(result)
```

---

## Integration Methods

### Method 1: Email-to-Task (Recommended)
- **Pros**: No API required, very reliable, works in restricted environments
- **Cons**: Cannot read existing tasks, one-way sync
- **Best for**: Creating tasks, bulk imports, automated workflows

### Method 2: CSV Import/Export
- **Pros**: Can handle 1000s of tasks at once, great for roadmaps
- **Cons**: Manual import step required in Asana UI
- **Best for**: Large bulk operations, project setup, quarterly planning

### Method 3: Browser Automation (Advanced)
- **Pros**: Can read and update tasks, full UI control
- **Cons**: Slower, UI changes can break automation
- **Best for**: Complex workflows needing bidirectional sync

---

## Customizing Workflows

All templates follow the same structure:

```python
class MyAsanaWorkflow(BaseModule):
    def configure(self, **kwargs):
        # Set up parameters
        pass

    def validate(self):
        # Validate inputs
        pass

    def execute(self):
        # Do the work
        from src.utils.workflow_helpers import AsanaHelper

        helper = AsanaHelper(project_email='...')
        result = helper.bulk_create_from_excel('file.xlsx')
        return result
```

### Adding Custom Task Templates

Edit `weekly_sprint_tasks.py` to add your own recurring tasks:

```python
task_templates = [
    {
        'name': 'My Custom Task',
        'assignee': 'john.doe',
        'day_offset': 2,  # Days from sprint start
        'description': 'Task description',
        'priority': 'High'
    }
]
```

---

## Troubleshooting

### Tasks not appearing in Asana
- **Check email address**: Verify you're using the correct project email
- **Check Outlook**: Ensure emails are being sent (check Sent folder)
- **Check Asana spam**: New project emails might be filtered

### Excel file not found
- Use absolute paths: `C:/Data/tasks.xlsx`
- Check file exists and is not open in Excel

### pywin32 not available
- Install: `pip install pywin32`
- Or use CSV method instead of email method

### Assignee not recognized
- Use exact Asana username (check in Asana profile)
- Don't include @ symbol in Excel file

---

## Advanced Usage

### Scheduling Weekly Sprint Tasks

Use the built-in scheduler to automatically create sprint tasks every Monday:

1. Open Automation Hub → Scheduler
2. Add new schedule
3. Select `weekly_sprint_tasks.py`
4. Set to run every Monday at 9 AM
5. Configure parameters

### Syncing Excel Tracker to Asana

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper(project_email='project@mail.asana.com')
result = helper.sync_excel_tracker_to_asana('C:/Trackers/project.xlsx')
```

### Parsing Asana Exports for Reporting

```python
from src.utils.workflow_helpers import AsanaHelper

helper = AsanaHelper()
tasks = helper.parse_asana_export('C:/Downloads/asana_export.csv')

# Analyze tasks
completed = [t for t in tasks if t['completed_at']]
print(f"Completed: {len(completed)}/{len(tasks)}")
```

---

## Support

For issues or questions:
1. Check the main documentation in `docs/`
2. Review example workflows in `user_scripts/templates/`
3. Use the AI Workflow Generator to create custom workflows

---

## Version History

- **1.0.0** (2024-11-10): Initial release
  - Excel to Asana bulk import
  - Weekly sprint task automation
  - CSV generation for manual import
  - Quarterly roadmap import
