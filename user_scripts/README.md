# User Scripts Directory

This directory contains custom automation workflows created by users.

## Directory Structure

- `custom/` - Your custom workflows created via the AI Workflow Generator
- `templates/` - Pre-built workflow templates that can be customized
  - `reports/` - Report generation workflows
  - `email/` - Email and communication workflows
  - `files/` - File and document management workflows

## Creating Custom Workflows

### Method 1: AI Workflow Generator (Recommended)
1. Open Automation Hub
2. Click "Create New Workflow" in the Scripts menu
3. Describe what you want in plain English
4. Review and customize the generated script
5. Save - it will automatically appear in your script library

### Method 2: Manual Script Creation
Create a Python file with the following structure:

```python
"""
WORKFLOW_META:
  name: My Custom Workflow
  description: What this workflow does
  category: Custom
  version: 1.0.0
  author: Your Name
  parameters:
    input_file:
      type: file
      description: Input file path
      required: true
    output_folder:
      type: string
      description: Where to save results
      required: true
      default: "C:/Output"
"""

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)

class CustomWorkflow(BaseModule):
    """Custom workflow implementation"""

    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_folder = None

    def configure(self, **kwargs):
        """Configure the workflow with parameters"""
        self.input_file = kwargs.get('input_file')
        self.output_folder = kwargs.get('output_folder')
        logger.info(f"Configured workflow: {self.input_file} -> {self.output_folder}")

    def validate(self):
        """Validate configuration"""
        if not self.input_file:
            raise ValueError("input_file is required")
        return True

    def execute(self):
        """Main workflow logic"""
        try:
            logger.info("Starting workflow execution...")

            # Your automation logic here

            logger.info("Workflow completed successfully!")
            return {"status": "success", "message": "Workflow completed"}
        except Exception as e:
            self.handle_error(e, "Workflow execution failed")
            raise

# Entry point for execution
def run(**kwargs):
    """Execute the workflow"""
    workflow = CustomWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
```

## Metadata Format

The WORKFLOW_META section supports:

- **name**: Display name in the GUI
- **description**: What the workflow does
- **category**: Category for organization (Custom, Reports, Email, Files, etc.)
- **version**: Version number
- **author**: Your name
- **parameters**: Dictionary of input parameters
  - **type**: string, text, file, choice, boolean
  - **description**: Help text
  - **required**: true/false
  - **default**: Default value
  - **choices**: List of options (for choice type)

## Best Practices

1. Always use the BaseModule class as foundation
2. Implement configure(), validate(), and execute() methods
3. Use the logger for output messages
4. Handle errors with try/except blocks
5. Return a result dictionary with status and details
6. Test workflows before scheduling them
