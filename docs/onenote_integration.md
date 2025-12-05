# OneNote Integration - Central Knowledge Repository

## Overview

The OneNote integration transforms your Automation Hub into a knowledge-powered automation platform. OneNote serves as your **central knowledge repository**, allowing you to:

- **Reference project facts** directly in automated reports
- **Automatically create** structured meeting minutes and documentation
- **Search and retrieve** critical information on-demand
- **Centralize research** from scattered files into organized notebooks

## Architecture

### Integration Approach: COM-Based (Corporate-Friendly)

The integration uses **COM automation** to connect with OneNote desktop application:

✅ **Works offline** - No internet connection required
✅ **No IT approval needed** - Uses your existing OneNote credentials
✅ **No API keys** - Direct application integration
✅ **Corporate-friendly** - Perfect for locked-down IT environments

### Components

```
automation_hub/src/modules/onenote/
├── com_client.py           # COM automation layer
├── content_formatter.py    # Rich content builder
├── note_manager.py         # High-level API (BaseModule)
└── __init__.py            # Module exports

automation_hub/src/utils/
└── workflow_helpers.py    # OneNoteHelper (AI workflows)

automation_hub/scripts/examples/
└── onenote_examples.py    # Example workflows
```

## Installation

### Prerequisites

1. **OneNote Desktop Application** - Must be installed and signed in
2. **pywin32 Package** - For COM automation (Windows only)

```bash
pip install pywin32
```

### Verification

Run the integration test:

```bash
python scripts/test_onenote_integration.py
```

Expected output:
```
🎉 All tests passed! OneNote integration is ready to use.
```

## Configuration

### GUI Configuration

1. **Open Automation Hub**
2. **Go to Settings** → OneNote tab
3. **Configure defaults:**
   - Default Notebook: e.g., "Project Knowledge"
   - Default Section: e.g., "Automation"
4. **Test Connection** to verify OneNote is accessible

### Programmatic Configuration

```python
from utils.workflow_helpers import OneNoteHelper

onenote = OneNoteHelper(
    default_notebook='Project Knowledge',
    default_section='Automation'
)
```

## Use Cases

### 1. Reference Data in Reports

**Extract project facts from OneNote and include them in Excel reports.**

```python
from utils.workflow_helpers import OneNoteHelper, ReportBuilder

# Extract facts from OneNote
onenote = OneNoteHelper()
facts = onenote.extract_facts_for_report(
    notebook='Project Knowledge',
    section='Metrics',
    page_title='Q4 Project Metrics'
)

# Create Excel report with OneNote data
report = ReportBuilder()
report.add_title_sheet("Q4 Status Report")

# Add tables extracted from OneNote
for table in facts['tables']:
    report.add_data_sheet("Metrics", data=table[1:], headers=table[0])

report.save('Q4_Status_Report.xlsx')
```

**Search and reference decisions:**

```python
# Search for approved decisions
decisions = onenote.search_for_facts(
    query='approved decision',
    notebook='Project Knowledge',
    max_results=5
)

# Add to report
for decision in decisions:
    print(f"{decision['title']}: {decision['snippet']}")
```

### 2. Automated Note Creation

**Automatically create structured meeting minutes from Outlook meetings.**

```python
from utils.workflow_helpers import OneNoteHelper

onenote = OneNoteHelper()

# Option A: From Outlook meeting
page_id = onenote.create_meeting_minutes_from_outlook(
    notebook='Meetings',
    section='Sprint Planning',
    meeting_subject='Sprint 42 Planning Meeting'
)

# Option B: Manual creation
page_id = onenote.create_meeting_minutes(
    notebook='Meetings',
    section='Sprint Planning',
    meeting_title='Sprint 42 Planning',
    date='2024-12-09',
    attendees=['Alice', 'Bob', 'Carol'],
    agenda=['Review backlog', 'Estimate tasks', 'Plan sprint'],
    action_items=[
        {'task': 'Update design docs', 'owner': 'Alice', 'due': '2024-12-12'}
    ]
)
```

**Create project status updates:**

```python
page_id = onenote.create_project_status_page(
    notebook='Projects',
    section='Q4 2024',
    project_name='Homepage Redesign',
    status='On Track',
    progress=75,
    milestones=[
        {'name': 'Design Complete', 'due': '2024-12-01', 'status': 'Done'},
        {'name': 'Development', 'due': '2024-12-15', 'status': 'In Progress'}
    ],
    issues=['Performance testing delayed by 2 days'],
    next_steps=['Complete development by Dec 15', 'Start QA testing']
)
```

### 3. Search and Retrieve Information

**Quick lookup of project requirements, decisions, or facts.**

```python
from utils.workflow_helpers import OneNoteHelper

onenote = OneNoteHelper()

# Search for requirements
results = onenote.search_for_facts(
    query='authentication requirements',
    notebook='Project Knowledge',
    max_results=10
)

for result in results:
    print(f"{result['title']}")
    print(f"Location: {result['notebook_name']} / {result['section_name']}")
    print(f"Snippet: {result['snippet']}")
    print()

# Get full content from a page
content = onenote.get_page_content(
    notebook='Project Knowledge',
    section='Requirements',
    title='Q4 Requirements',
    format='text'
)
```

**Search and email results to stakeholders:**

```python
from utils.workflow_helpers import OneNoteHelper, EmailHelper

# Search OneNote
onenote = OneNoteHelper()
decisions = onenote.search_for_facts(
    query='approved',
    notebook='Project Knowledge',
    max_results=5
)

# Email results
email = EmailHelper()
body = "<h2>Recent Approved Decisions</h2><ul>"
for decision in decisions:
    body += f"<li><b>{decision['title']}</b>: {decision['snippet']}</li>"
body += "</ul>"

email.send_report(
    to=['stakeholders@company.com'],
    subject='Weekly Decisions Update',
    body=body
)
```

### 4. Centralize Research Findings

**Aggregate scattered research files into organized OneNote pages.**

```python
from utils.workflow_helpers import OneNoteHelper

onenote = OneNoteHelper()

# Aggregate multiple research files
page_id = onenote.aggregate_research_to_central_notebook(
    source_files=[
        'C:/Research/competitor_analysis.txt',
        'C:/Research/user_research_notes.txt',
        'C:/Research/tech_stack_research.txt'
    ],
    central_notebook='Project Knowledge',
    section='Research',
    topic='Homepage Redesign - Comprehensive Research'
)
```

**Create standalone research notes:**

```python
page_id = onenote.create_research_notes(
    notebook='Research',
    section='Competitors',
    topic='Q4 2024 Competitor Homepage Analysis',
    source='Web research and analytics data - December 2024',
    key_findings=[
        'Average page load time: 2.3 seconds',
        'Mobile traffic: 65% of visits',
        '80% feature video content above the fold',
        'Average conversion rate: 3.2%'
    ],
    details="""
    Detailed Analysis:
    - Performance: Fastest 1.2s, slowest 4.1s, median 2.3s
    - Design trends: Minimalist with bold typography (60%)
    - Content strategy: Hero section with CTA (100%)
    """,
    tags=['homepage', 'competitors', 'q4-2024', 'analysis']
)
```

## Advanced Features

### Custom Content Builder

Create rich, structured OneNote content programmatically:

```python
from modules.onenote import OneNoteContentBuilder

builder = OneNoteContentBuilder("Project Status Report")

# Add structured content
builder.add_heading("Executive Summary", 1)
builder.add_text("Project is on track for Q4 delivery.", bold=True)
builder.add_blank_line()

builder.add_heading("Key Metrics", 2)
builder.add_table(
    headers=["Metric", "Target", "Actual", "Status"],
    rows=[
        ["Load Time", "< 2s", "1.8s", "✓"],
        ["Conversion", "3.5%", "3.2%", "⚠"],
        ["Mobile Traffic", "60%", "65%", "✓"]
    ]
)

builder.add_heading("Next Steps", 2)
builder.add_numbered_list([
    "Optimize conversion funnel",
    "Complete performance testing",
    "Deploy to production"
])

# Build content
content = builder.build_simple()

# Create page
onenote = OneNoteHelper()
page_id = onenote.create_page(
    notebook='Projects',
    section='Q4 2024',
    title='Weekly Status Report',
    content=content
)
```

### Template System

Pre-built templates for common scenarios:

```python
from modules.onenote import TemplateBuilder

# Meeting minutes
template = TemplateBuilder.meeting_minutes(
    meeting_title="Sprint Planning",
    date="2024-12-09",
    attendees=["Team Lead", "Dev 1", "Dev 2"],
    agenda_items=["Review sprint", "Plan next sprint"],
    action_items=[{"task": "Task 1", "owner": "Dev 1", "due": "2024-12-15"}]
)

# Project status
template = TemplateBuilder.project_status(
    project_name="Homepage Redesign",
    status="On Track",
    progress=75,
    milestones=[{"name": "Milestone 1", "due": "2024-12-15", "status": "Done"}]
)

# Research notes
template = TemplateBuilder.research_notes(
    topic="Competitor Analysis",
    source="Web research",
    key_findings=["Finding 1", "Finding 2"],
    tags=["research", "competitors"]
)

# Decision log
template = TemplateBuilder.decision_log(
    decision_title="Technology Stack Choice",
    context="Need to choose frontend framework",
    options=[
        {"name": "React", "pros": "Popular", "cons": "Complex"},
        {"name": "Vue", "pros": "Simple", "cons": "Smaller ecosystem"}
    ],
    chosen_option="React",
    rationale="Better long-term support",
    stakeholders=["Tech Lead", "CTO"]
)
```

### Notebook Management

List and explore your OneNote structure:

```python
onenote = OneNoteHelper()

# List all notebooks
notebooks = onenote.list_notebooks()
for nb in notebooks:
    print(f"Notebook: {nb['name']}")

# List sections in a notebook
sections = onenote.list_sections('Project Knowledge')
for sec in sections:
    print(f"  Section: {sec['name']}")

# List pages in a section
pages = onenote.list_pages('Project Knowledge', 'Requirements')
for page in pages:
    print(f"    Page: {page['title']} ({page['last_modified']})")
```

## Complete Workflow Example

**Weekly project status workflow combining all use cases:**

```python
from utils.workflow_helpers import OneNoteHelper, ReportBuilder, EmailHelper
from datetime import datetime

onenote = OneNoteHelper()

# 1. Aggregate weekly research (Use Case 4)
# (Research files collected throughout the week)

# 2. Create meeting minutes (Use Case 2)
meeting_page_id = onenote.create_meeting_minutes(
    notebook='Meetings',
    section='Weekly Status',
    meeting_title=f'Weekly Status - {datetime.now().strftime("%Y-%m-%d")}',
    date=datetime.now().strftime('%Y-%m-%d'),
    attendees=['Product Owner', 'Tech Lead', 'Design Lead'],
    agenda=['Review progress', 'Discuss blockers', 'Plan next week'],
    action_items=[
        {'task': 'Complete user testing', 'owner': 'Design Lead', 'due': '2024-12-15'}
    ]
)

# 3. Search for key decisions (Use Case 3)
decisions = onenote.search_for_facts(
    query='approved',
    notebook='Project Knowledge',
    max_results=5
)

# 4. Generate status report with OneNote data (Use Case 1)
report = ReportBuilder()
report.add_title_sheet(
    title="Weekly Status Report",
    subtitle=datetime.now().strftime("%Y-%m-%d"),
    metadata={
        "Meeting Minutes": f"Page ID: {meeting_page_id}",
        "Key Decisions": f"{len(decisions)} referenced"
    }
)

# Add decisions to report
decision_data = [[d['title'], d['snippet'][:100]] for d in decisions]
report.add_data_sheet("Key Decisions", headers=["Decision", "Summary"], data=decision_data)

report.save('Weekly_Status_Complete.xlsx')

# Email report
email = EmailHelper()
email.send_report(
    to=['stakeholders@company.com'],
    subject='Weekly Status Report',
    body='Please see attached weekly status report.',
    attachments=['Weekly_Status_Complete.xlsx']
)

print("✓ Weekly workflow complete!")
print("✓ OneNote is your central knowledge repository")
```

## API Reference

### OneNoteHelper

Main class for OneNote automation workflows.

#### Constructor

```python
OneNoteHelper(default_notebook=None, default_section=None)
```

#### Methods

**Page Operations:**
- `create_page(notebook, section, title, content="")`
- `get_page_content(notebook, section, title, format="text")`
- `update_page(notebook, section, title, content, new_title=None)`
- `delete_page(notebook, section, title)`

**Search:**
- `search_for_facts(query, notebook=None, max_results=10)`

**Data Extraction:**
- `extract_facts_for_report(notebook, section, page_title)`

**Templates:**
- `create_meeting_minutes(notebook, section, meeting_title, date, attendees, agenda, notes="", action_items=None)`
- `create_meeting_minutes_from_outlook(notebook, section, meeting_subject)`
- `create_project_status_page(notebook, section, project_name, status, progress, milestones, issues=None, next_steps=None)`
- `create_research_notes(notebook, section, topic, source, key_findings, details="", tags=None)`

**Aggregation:**
- `aggregate_research_to_central_notebook(source_files, central_notebook, section, topic)`

**Navigation:**
- `list_notebooks()`
- `list_sections(notebook)`
- `list_pages(notebook, section)`

### OneNoteManager

Lower-level API inheriting from BaseModule.

```python
from modules.onenote import OneNoteManager

manager = OneNoteManager()
manager.configure(default_notebook='Project Knowledge', default_section='Automation')
manager.validate()  # Connects to OneNote

# Use manager methods
notebooks = manager.list_notebooks()
page_id = manager.create_page(notebook, section, title, content)
```

### OneNoteContentBuilder

Rich content builder for creating structured OneNote pages.

```python
from modules.onenote import OneNoteContentBuilder

builder = OneNoteContentBuilder("Page Title")
builder.add_heading("Heading", level=1)
builder.add_text("Text", bold=False, italic=False)
builder.add_bullet_list(["Item 1", "Item 2"])
builder.add_numbered_list(["Step 1", "Step 2"])
builder.add_table(headers=["H1", "H2"], rows=[["R1C1", "R1C2"]])
builder.add_code_block("code", language="python")
builder.add_divider()
builder.add_blank_line()

content = builder.build_simple()  # Returns text
# or
xml = builder.build_xml()  # Returns OneNote XML
```

## Troubleshooting

### Connection Issues

**Problem:** "OneNote desktop application not found"

**Solution:**
1. Install OneNote desktop application
2. Sign in with your Microsoft account
3. Ensure OneNote is running or has been launched at least once

**Problem:** "pywin32 is required"

**Solution:**
```bash
pip install pywin32
```

After installation, you may need to run:
```bash
python Scripts/pywin32_postinstall.py -install
```

### Permission Issues

**Problem:** "Access denied" when creating pages

**Solution:**
1. Ensure OneNote is signed in
2. Verify notebooks exist in OneNote
3. Check that sections are not password-protected

### Performance

**Problem:** Slow search or large notebook performance

**Solution:**
- Limit search results with `max_results` parameter
- Create separate notebooks for different projects
- Archive old pages regularly

## Best Practices

### Organization

1. **Create dedicated notebooks:**
   - "Project Knowledge" for project information
   - "Meetings" for meeting minutes
   - "Research" for research notes

2. **Use consistent section names:**
   - "Requirements" for requirements
   - "Decisions" for decision logs
   - "Status" for status updates

3. **Tag pages appropriately:**
   - Use tags for categorization
   - Include dates in titles
   - Use consistent naming conventions

### Automation

1. **Configure defaults in GUI** for consistent automation
2. **Use templates** for standardized content
3. **Search before creating** to avoid duplicates
4. **Aggregate regularly** to centralize knowledge

### Integration

1. **Combine with Excel** for reports with OneNote references
2. **Link with Outlook** for meeting minutes automation
3. **Use with SharePoint** for document backup workflows
4. **Integrate with Asana** for project tracking with OneNote facts

## Examples

See comprehensive examples in:
- `/scripts/examples/onenote_examples.py`

Run examples:
```bash
python scripts/examples/onenote_examples.py
```

## Future Enhancements (Optional)

While the current COM-based implementation meets all requirements for locked-down environments, future enhancements could include:

- **Microsoft Graph API support** (for cloud-based scenarios with IT approval)
- **Image insertion** (screenshots, diagrams)
- **Advanced table formatting** (merged cells, styling)
- **Section creation** (currently requires manual creation)
- **Notebook creation** (currently requires manual creation)
- **OneNote tags** (todo, important, question)

These would require additional setup and potentially IT approval but would provide enhanced capabilities.

## Support

For issues or questions:
1. Check troubleshooting section above
2. Review example workflows
3. Run integration test: `python scripts/test_onenote_integration.py`
4. Check application logs in `/logs/` directory

---

**OneNote integration complete! Your central knowledge repository is ready.**
