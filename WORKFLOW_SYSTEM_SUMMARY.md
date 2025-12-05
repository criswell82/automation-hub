# Automation Hub - AI Workflow Generation System

## 🎉 Implementation Complete!

Your Automation Hub has been enhanced with a powerful AI-assisted workflow creation system tailored to your specific needs: **report generation**, **email/communication**, and **file management**.

---

## ✨ What's New

### 1. **AI Workflow Generator**
Create custom automation workflows by describing what you want in plain English. The system uses Claude AI to generate Python scripts automatically.

**How to Use:**
- Open Automation Hub
- Click `File` → `Create New Workflow` (or press `Ctrl+N`)
- Describe what you want to automate
- Review and save the generated code
- Workflow appears instantly in your script library!

### 2. **Dynamic Script Loading System**
No more hardcoded scripts! The system now automatically discovers and loads custom workflows from the `user_scripts/` directory.

**Features:**
- Auto-discovery of custom scripts
- Hot-reload capability (refresh without restarting)
- Supports custom categories
- Dynamic parameter generation

### 3. **Workflow Helper Library**
Pre-built helper functions for common automation patterns:

- **ReportBuilder**: Create Excel reports with charts
- **EmailHelper**: Send emails, process inbox, extract attachments
- **FileOrganizer**: Organize by type/date, archive old files
- **SharePointHelper**: Sync folders to SharePoint

### 4. **Pre-Built Templates**
Three ready-to-use workflow templates for your specific needs:

1. **Weekly Report Generator** - Generate and email Excel reports from data files
2. **Email Attachment Processor** - Extract and organize email attachments
3. **File Organizer & SharePoint Sync** - Organize files and sync to SharePoint

---

## 📁 Project Structure

```
automation_hub/
├── src/
│   ├── core/
│   │   ├── script_discovery.py          # NEW: Dynamic script loader
│   │   └── ai_workflow_generator.py     # NEW: AI generation backend
│   ├── hub/
│   │   ├── workflow_generator_dialog.py # NEW: AI generator UI
│   │   ├── script_manager.py            # UPDATED: Dynamic loading
│   │   └── main_window.py               # UPDATED: Menu integration
│   └── utils/
│       └── workflow_helpers.py          # NEW: Helper functions library
│
├── user_scripts/                         # NEW: Custom workflows directory
│   ├── custom/                           # Your generated workflows
│   │   ├── example_hello_world.py       # Example custom workflow
│   │   └── [your workflows here]
│   └── templates/                        # Pre-built templates
│       ├── reports/
│       │   └── weekly_report_generator.py
│       ├── email/
│       │   └── email_attachment_processor.py
│       └── files/
│           └── file_organizer_and_sync.py
│
├── docs/
│   └── AI_Workflow_Generator_Guide.md   # NEW: Comprehensive guide
│
└── requirements.txt                      # UPDATED: Added anthropic package
```

---

## 🚀 Quick Start Guide

### Step 1: Install AI Dependencies (Optional)

For AI-powered workflow generation:

```bash
pip install anthropic
```

Set your API key:
```bash
# Windows Command Prompt
set ANTHROPIC_API_KEY=your_api_key_here

# Or add to system environment variables
```

**Note:** The system works without an API key using template-based generation.

### Step 2: Launch Automation Hub

```bash
cd automation_hub
python src/main.py
```

### Step 3: Create Your First Workflow

1. Click **File** → **Create New Workflow** (Ctrl+N)
2. Describe your automation need, for example:

   ```
   Generate an Excel report from CSV files in C:/Data,
   add a summary chart, and email it to manager@company.com
   ```

3. Click **Generate Workflow**
4. Review the generated code
5. Click **Save Workflow**
6. The workflow appears in your script library!

### Step 4: Try Pre-Built Templates

The templates are already loaded! Look for:
- "Weekly Report Generator" (Reports category)
- "Email Attachment Processor" (Email category)
- "File Organizer and SharePoint Sync" (Files category)

Select one, configure parameters, and run!

---

## 💡 Common Use Cases

### 1. Generate Weekly Sales Report

**What you want:** Compile weekly sales data from multiple CSV files into an Excel report with charts and email it to your team.

**Solution:**
1. Use the "Weekly Report Generator" template, or
2. Generate custom workflow:
   ```
   Process all CSV files in C:/Sales/Weekly, create an Excel report
   with sales summary and charts, email to team@company.com every Friday at 5 PM
   ```

### 2. Process Invoice Emails

**What you want:** Automatically extract PDF invoices from emails and organize them by month.

**Solution:**
1. Use the "Email Attachment Processor" template, or
2. Generate custom workflow:
   ```
   Monitor inbox for emails with "Invoice" in subject, extract PDF attachments,
   save to C:/Invoices organized by month, mark emails as read
   ```

### 3. Organize and Backup Files

**What you want:** Keep your Downloads folder organized and sync important files to SharePoint.

**Solution:**
1. Use the "File Organizer & SharePoint Sync" template, or
2. Generate custom workflow:
   ```
   Organize Downloads folder by file type, archive files older than 90 days,
   sync documents to SharePoint at https://company.sharepoint.com/sites/docs
   ```

---

## 🎯 Key Features

### AI Generation
- **Natural language input**: Describe in plain English
- **Automatic code generation**: Creates complete Python scripts
- **Smart parameter detection**: Extracts parameters from description
- **Context-aware**: Understands your automation modules

### Template System
- **Pre-built templates**: Ready for common tasks
- **Customizable**: Modify to match your exact needs
- **Well-documented**: Comments explain each step
- **Tested**: Working examples included

### Helper Library
- **Report generation**: Excel reports with charts
- **Email automation**: Send, receive, process attachments
- **File management**: Organize, rename, archive
- **SharePoint integration**: Upload, download, sync

### Dynamic Loading
- **Auto-discovery**: Scripts automatically appear in GUI
- **Hot-reload**: Refresh without restarting (F5)
- **Custom categories**: Organize your workflows
- **Parameter UI**: Automatic form generation

---

## 📚 Documentation

### Main Guides
- **AI_Workflow_Generator_Guide.md** - Complete user guide for the AI system
- **user_scripts/README.md** - Guide for creating custom workflows manually
- **Original documentation** - Still available in `docs/` folder

### Examples
- **user_scripts/custom/example_hello_world.py** - Simple example
- **user_scripts/templates/** - Three complete template workflows

---

## 🔧 Technical Details

### Architecture

**Phase 1: Dynamic Script Loading**
- `script_discovery.py`: Scans user_scripts/ for Python files
- Parses WORKFLOW_META from docstrings (YAML format)
- Dynamically imports and executes custom workflows
- Integrated with existing ScriptManager

**Phase 2: AI Workflow Generator**
- `workflow_generator_dialog.py`: PyQt5 UI for workflow creation
- `ai_workflow_generator.py`: Claude API integration
- Template-based fallback when API unavailable
- Syntax highlighting in code preview

**Phase 3: Helper Library**
- `workflow_helpers.py`: Reusable automation components
- Focused on report generation, email, and file management
- Built on top of existing automation modules
- Common workflow patterns implemented

**Phase 4: Integration & Polish**
- Menu integration in main window
- Auto-refresh after workflow creation
- Comprehensive documentation
- Example workflows and templates

### Workflow Metadata Format

```python
"""
WORKFLOW_META:
  name: Workflow Name
  description: What it does
  category: Reports/Email/Files/Custom
  version: 1.0.0
  author: Author Name
  parameters:
    param_name:
      type: string/text/file/choice/boolean
      description: Help text
      required: true/false
      default: "default value"
      choices: [option1, option2]  # For choice type
"""
```

### Supported Parameter Types
- **string**: Single-line text input (QLineEdit)
- **text**: Multi-line text input (QTextEdit)
- **file**: File picker dialog (QLineEdit + Browse button)
- **choice**: Dropdown selection (QComboBox)
- **boolean**: Checkbox (QCheckBox)

---

## 🛠️ Installation & Setup

### Install Dependencies

```bash
cd automation_hub
pip install -r requirements.txt
```

### Optional: AI Generation

```bash
pip install anthropic
```

Get your API key from: https://console.anthropic.com/

### Run the Application

```bash
python src/main.py
```

---

## 📝 Creating Workflows Manually

If you prefer to write workflows yourself:

1. Create a Python file in `user_scripts/custom/`
2. Add WORKFLOW_META docstring with parameters
3. Inherit from BaseModule
4. Implement configure(), validate(), and execute() methods
5. Add a run(**kwargs) function
6. Refresh scripts in the GUI (F5)

See `user_scripts/README.md` for detailed instructions.

---

## 🎨 Customization Examples

### Modify a Template

1. Open template file (e.g., `weekly_report_generator.py`)
2. Modify the `execute()` method logic
3. Add custom data processing
4. Save to `user_scripts/custom/` with a new name
5. Refresh scripts (F5)

### Add Custom Parameters

```python
"""
WORKFLOW_META:
  parameters:
    custom_param:
      type: choice
      choices: ['option1', 'option2', 'option3']
      description: Your custom parameter
      required: true
      default: 'option1'
"""
```

### Use Helper Functions

```python
from src.utils.workflow_helpers import ReportBuilder, EmailHelper

# In your execute() method:
report = ReportBuilder()
report.add_title_sheet("My Report", "Subtitle")
# ... add data ...
report.save("output.xlsx")

email = EmailHelper()
email.send_report(['user@example.com'], 'Report', 'Body', ['output.xlsx'])
```

---

## 🐛 Troubleshooting

### Workflow Doesn't Appear in Library
- Click **Scripts** → **Refresh Scripts** (F5)
- Check file is in `user_scripts/custom/` or `user_scripts/templates/`
- Verify WORKFLOW_META is properly formatted

### AI Generation Fails
- Check API key is set correctly
- Verify internet connection
- System will fall back to templates if AI unavailable

### Execution Errors
- Review logs in the Output tab
- Verify file paths exist
- Check parameter values are correct
- Ensure required modules are installed

### Import Errors
- Run `pip install -r requirements.txt`
- For AI: `pip install anthropic`
- Check Python version (requires 3.8+)

---

## 🚀 Next Steps

1. **Test the System**
   - Try the AI workflow generator
   - Run the example workflows
   - Explore the templates

2. **Create Your Workflows**
   - Generate workflows for your specific tasks
   - Customize templates to match your needs
   - Save commonly-used workflows

3. **Schedule Automation**
   - Set up recurring workflows
   - Schedule reports to run automatically
   - Automate repetitive tasks

4. **Expand & Customize**
   - Add custom helper functions
   - Create your own templates
   - Share workflows with colleagues (future feature)

---

## 📊 What You Get

### ✅ Completed Implementation

- [x] Dynamic script loading system
- [x] AI workflow generator UI
- [x] Claude AI integration
- [x] Template-based fallback
- [x] Workflow helper library
- [x] 3 pre-built templates for your use cases
- [x] Menu integration
- [x] Auto-refresh capability
- [x] Comprehensive documentation
- [x] Example workflows

### 🎯 Tailored to Your Needs

- **Report Generation**: Excel reports with data aggregation and charts
- **Email & Communication**: Process emails, extract attachments, send reports
- **File Management**: Organize files, archive old data, sync to SharePoint

### 🧰 Tools Provided

- **ReportBuilder**: Build Excel reports easily
- **EmailHelper**: Email automation simplified
- **FileOrganizer**: File management operations
- **SharePointHelper**: SharePoint integration
- **Script Discovery**: Automatic workflow loading
- **AI Generator**: Natural language workflow creation

---

## 💬 Usage Tips

1. **Start Simple**: Begin with pre-built templates
2. **Test First**: Run workflows manually before scheduling
3. **Use Descriptive Names**: Clear names make workflows easy to find
4. **Document Changes**: Add comments when customizing
5. **Backup Workflows**: Keep copies of important workflows
6. **Schedule Wisely**: Test timing and frequency before production use

---

## 🎓 Learning Resources

- **AI_Workflow_Generator_Guide.md**: Complete guide to using the system
- **user_scripts/README.md**: Manual workflow creation guide
- **Template workflows**: Three working examples with comments
- **Example workflow**: Simple hello world to understand structure
- **Helper library source**: Check `workflow_helpers.py` for all available functions

---

## 🏆 Summary

You now have a complete AI-assisted workflow creation system that allows you to:

1. **Generate workflows from natural language** using Claude AI
2. **Use pre-built templates** for common tasks (reports, email, files)
3. **Leverage helper functions** for rapid development
4. **Customize and extend** workflows to match your exact needs
5. **Schedule automation** for recurring tasks
6. **Integrate seamlessly** with existing Automation Hub features

The system is designed to minimize coding time while maximizing automation capability - exactly what you requested!

---

## 📞 Getting Started

1. Launch Automation Hub: `python src/main.py`
2. Press `Ctrl+N` to open the workflow generator
3. Describe what you want to automate
4. Let AI generate the code
5. Save and run!

**Happy Automating! 🚀**
