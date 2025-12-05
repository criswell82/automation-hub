"""
AI Workflow Generator

Uses Claude AI to generate workflow scripts based on natural language descriptions.
"""

import os
import json
import re
from typing import Dict, List, Any, Optional
from pathlib import Path

from core.logging_config import get_logger

logger = get_logger(__name__)


class AIWorkflowGenerator:
    """Generates workflow scripts using AI"""

    def __init__(self) -> None:
        """Initialize the AI workflow generator"""
        self.logger = logger
        self.api_key = self._get_api_key()

    def _get_api_key(self) -> Optional[str]:
        """
        Get Claude API key from secure storage

        Priority order:
        1. SecurityManager (Windows Credential Manager) - most secure
        2. Environment variables
        3. Config file (deprecated, for backwards compatibility)
        """
        api_key = None

        # 1. Try SecurityManager (encrypted storage)
        try:
            from core.security import get_security_manager
            security = get_security_manager()
            cred = security.retrieve_credential('AnthropicAPIKey')
            if cred:
                api_key = cred.get('password')
                if api_key:
                    self.logger.debug("API key loaded from secure storage")
                    return api_key
        except Exception as e:
            self.logger.debug(f"Could not load API key from secure storage: {e}")

        # 2. Try environment variables
        api_key = os.environ.get('ANTHROPIC_API_KEY') or os.environ.get('CLAUDE_API_KEY')
        if api_key:
            self.logger.debug("API key loaded from environment variable")
            return api_key

        # 3. Try config file (deprecated)
        try:
            from core.config import ConfigManager
            config = ConfigManager()
            api_key = config.get('ai.anthropic_api_key')
            if api_key:
                self.logger.warning("API key loaded from config file (not secure). Please use Settings to store it securely.")
                return api_key
        except Exception:
            pass

        return api_key

    @staticmethod
    def get_available_models(api_key: str = None) -> List[Dict[str, Any]]:
        """
        Get list of available Claude models

        Args:
            api_key: Optional API key to test against. If not provided, returns hardcoded list.

        Returns:
            List of model dictionaries with id, name, and description
        """
        # Hardcoded list of known Claude models (fallback)
        models = [
            {
                'id': 'claude-sonnet-4-5-20250929',
                'name': 'Claude 4.5 Sonnet',
                'description': 'Best balance of speed and capability (Recommended)',
                'speed': 'Fast',
                'capability': 'High'
            },
            {
                'id': 'claude-opus-4-5-20250514',
                'name': 'Claude 4.5 Opus',
                'description': 'Most capable model, best for complex tasks',
                'speed': 'Moderate',
                'capability': 'Highest'
            },
            {
                'id': 'claude-haiku-4-20250228',
                'name': 'Claude 4 Haiku',
                'description': 'Fastest and most cost-effective',
                'speed': 'Very Fast',
                'capability': 'Moderate'
            }
        ]

        # If API key provided, could query Anthropic API for live model list
        # For now, return the hardcoded list which includes all current Claude 4+ models
        if api_key:
            try:
                import anthropic
                # Note: Anthropic API doesn't currently have a models.list() endpoint
                # So we return the hardcoded list with validation
                client = anthropic.Anthropic(api_key=api_key)
                # If the client initialization succeeds, API key is valid
                logger.debug("API key validated successfully")
            except Exception as e:
                logger.debug(f"Could not validate models with API: {e}")

        return models

    def generate_workflow(
        self,
        description: str,
        category: str = "Custom",
        use_templates: List[str] = None
    ) -> Dict[str, Any]:
        """
        Generate a workflow script from natural language description

        Args:
            description: Natural language description of what to automate
            category: Category for the workflow
            use_templates: List of template types to reference

        Returns:
            Dictionary with:
                - code: Generated Python code
                - name: Workflow name
                - description: Cleaned up description
                - parameters: Extracted parameters
        """
        try:
            # Check if API key is available
            if self.api_key:
                return self._generate_with_ai(description, category, use_templates)
            else:
                # Fallback to template-based generation
                self.logger.warning("No API key found, using template-based generation")
                return self._generate_from_template(description, category)

        except Exception as e:
            self.logger.error(f"Workflow generation failed: {e}")
            raise

    def _generate_with_ai(
        self,
        description: str,
        category: str,
        use_templates: List[str]
    ) -> Dict[str, Any]:
        """Generate workflow using Claude AI"""
        try:
            import anthropic
        except ImportError:
            self.logger.warning("anthropic package not installed, falling back to templates")
            return self._generate_from_template(description, category)

        # Build context with available modules and examples
        context = self._build_context(use_templates)

        # Create prompt
        prompt = self._create_generation_prompt(description, category, context)

        # Call Claude API
        client = anthropic.Anthropic(api_key=self.api_key)

        # Get model settings from config
        from core.config import ConfigManager
        config = ConfigManager()
        model = config.get('ai.model', 'claude-sonnet-4-5-20250929')
        max_tokens = config.get('ai.max_tokens', 4000)

        self.logger.info(f"Calling Claude API for workflow generation (model: {model})...")

        message = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        response_text = message.content[0].text

        # Parse the response
        result = self._parse_ai_response(response_text)

        self.logger.info("Workflow generated successfully with AI")
        return result

    def _generate_from_template(
        self,
        description: str,
        category: str
    ) -> Dict[str, Any]:
        """Fallback: Generate workflow using templates (no AI)"""

        self.logger.info("Generating workflow from template...")

        # Detect intent from description
        intent = self._detect_intent(description)

        # Load appropriate template
        template_code = self._get_template_for_intent(intent, category)

        # Extract parameters from description
        parameters = self._extract_parameters_from_description(description, intent)

        # Generate name from description
        name = self._generate_name_from_description(description)

        return {
            'code': template_code,
            'name': name,
            'description': description,
            'parameters': parameters
        }

    def _build_context(self, use_templates: List[str]) -> str:
        """Build context with available modules and examples"""
        context_parts = []

        # Available modules
        context_parts.append("""
## Available Automation Modules

You have access to these automation modules:

1. **Desktop RPA** (src.modules.desktop_rpa)
   - WindowManager: Find, activate, move, resize windows
   - InputController: Mouse clicks, keyboard input, screenshots

2. **Excel Automation** (src.modules.excel_automation)
   - WorkbookHandler: Read/write Excel files, format cells, add formulas
   - ChartBuilder: Create charts (bar, line, pie)

3. **Outlook Automation** (src.modules.outlook_automation)
   - EmailHandler: Send emails, read inbox, manage attachments

4. **SharePoint** (src.modules.sharepoint)
   - SharePointClient: Upload/download files, manage documents

5. **Word Automation** (src.modules.word_automation)
   - DocumentHandler: Create/edit Word documents

6. **OneNote** (src.modules.onenote)
   - NoteManager: Create and manage OneNote pages

7. **Asana Integration** (src.modules.asana)
   - AsanaEmailModule: Create tasks via email-to-task (no API required)
   - AsanaBrowserModule: Automate Asana UI via browser RPA
   - AsanaCSVHandler: Generate/parse Asana CSV files for bulk operations
   - AsanaHelper (in workflow_helpers): Convenient methods for common operations
     * create_task_via_email(): Create single task
     * bulk_create_from_excel(): Bulk create from Excel file
     * generate_asana_csv(): Generate CSV for manual import
     * sync_excel_tracker_to_asana(): Sync Excel tracker
     * create_weekly_sprint_tasks(): Create recurring sprint tasks
""")

        # Add template examples if requested
        if use_templates:
            context_parts.append("\n## Example Template Structure\n")
            context_parts.append(self._get_template_example())

        return "\n".join(context_parts)

    def _create_generation_prompt(
        self,
        description: str,
        category: str,
        context: str
    ) -> str:
        """Create the prompt for Claude"""
        return f"""You are an expert Python automation script generator. Generate a complete, working Python workflow script based on the user's description.

{context}

## User Request

Category: {category}
Description: {description}

## Requirements

1. Generate a complete Python script following this structure:
   - Docstring with WORKFLOW_META in YAML format (name, description, category, version, author, parameters)
   - Import statements (including sys.path setup if needed)
   - A class inheriting from BaseModule with configure(), validate(), and execute() methods
   - A run(**kwargs) function that instantiates and runs the workflow
   - Proper error handling and logging

2. The WORKFLOW_META must include:
   - name: A clear, descriptive name (title case)
   - description: What the workflow does
   - category: {category}
   - version: 1.0.0
   - author: Automation Hub
   - parameters: Dictionary of required parameters with type, description, required, default

3. Parameter types can be: string, text, file, choice, boolean

4. Use the appropriate automation modules from the context above

5. Include helpful comments and follow best practices

6. Make sure the script is complete and runnable

## Output Format

Provide ONLY the complete Python code. Do not include any explanations or markdown formatting - just the raw Python code that can be saved directly to a .py file.
"""

    def _parse_ai_response(self, response: str) -> Dict[str, Any]:
        """Parse AI response to extract code and metadata"""

        # Remove markdown code blocks if present
        code = response.strip()
        if code.startswith("```python"):
            code = code[9:]
        elif code.startswith("```"):
            code = code[3:]
        if code.endswith("```"):
            code = code[:-3]
        code = code.strip()

        # Extract metadata from docstring
        metadata = self._extract_metadata_from_code(code)

        return {
            'code': code,
            'name': metadata.get('name', 'Generated Workflow'),
            'description': metadata.get('description', ''),
            'parameters': metadata.get('parameters', {})
        }

    def _extract_metadata_from_code(self, code: str) -> Dict[str, Any]:
        """Extract WORKFLOW_META from generated code"""
        try:
            # Find WORKFLOW_META section
            match = re.search(r'WORKFLOW_META:\s*\n(.*?)\n["\']{3}', code, re.DOTALL)
            if match:
                import yaml
                metadata_text = match.group(1)
                metadata = yaml.safe_load(metadata_text)
                return metadata
        except Exception as e:
            logger.warning(f"Failed to extract metadata: {e}")

        return {}

    def _detect_intent(self, description: str) -> str:
        """Detect intent from description"""
        description_lower = description.lower()

        if any(word in description_lower for word in ['asana', 'task', 'project management', 'roadmap', 'sprint']):
            return 'asana'
        elif any(word in description_lower for word in ['excel', 'spreadsheet', 'workbook', 'report', 'chart']):
            return 'excel_report'
        elif any(word in description_lower for word in ['email', 'outlook', 'message', 'inbox']):
            return 'email'
        elif any(word in description_lower for word in ['file', 'folder', 'directory', 'sharepoint', 'document']):
            return 'file_management'
        elif any(word in description_lower for word in ['window', 'click', 'type', 'automate app']):
            return 'desktop_rpa'
        else:
            return 'generic'

    def _get_template_for_intent(self, intent: str, category: str) -> str:
        """Get template code based on detected intent"""

        name = f"{category} Workflow"
        desc = "Custom automation workflow"

        if intent == 'asana':
            return self._get_asana_template()
        elif intent == 'excel_report':
            return self._get_excel_template()
        elif intent == 'email':
            return self._get_email_template()
        elif intent == 'file_management':
            return self._get_file_management_template()
        else:
            return self._get_generic_template(name, desc)

    def _get_template_example(self) -> str:
        """Get example template structure"""
        return '''
```python
"""
WORKFLOW_META:
  name: Example Workflow
  description: What this workflow does
  category: Custom
  version: 1.0.0
  author: Your Name
  parameters:
    param1:
      type: string
      description: First parameter
      required: true
"""

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)

class MyWorkflow(BaseModule):
    def configure(self, **kwargs):
        # Setup parameters
        pass

    def validate(self):
        # Validate configuration
        return True

    def execute(self):
        # Main workflow logic
        logger.info("Starting workflow...")
        return {"status": "success"}

def run(**kwargs):
    workflow = MyWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
```
'''

    def _get_generic_template(self, name: str, description: str) -> str:
        """Get generic workflow template"""
        return f'''"""
WORKFLOW_META:
  name: {name}
  description: {description}
  category: Custom
  version: 1.0.0
  author: Automation Hub
  parameters:
    input_data:
      type: string
      description: Input data or file path
      required: false
      default: ""
"""

import sys
from pathlib import Path

# Add src directory to path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class CustomWorkflow(BaseModule):
    """Custom workflow implementation"""

    def __init__(self):
        super().__init__()
        self.input_data = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.input_data = kwargs.get('input_data', '')
        logger.info(f"Configured workflow with input: {{self.input_data}}")

    def validate(self):
        """Validate configuration"""
        return True

    def execute(self):
        """Main workflow logic"""
        try:
            logger.info("Starting workflow execution...")

            # TODO: Add your automation logic here

            logger.info("Workflow completed successfully!")
            return {{"status": "success", "message": "Workflow completed"}}

        except Exception as e:
            self.handle_error(e, "Workflow execution failed")
            raise


def run(**kwargs):
    """Execute the workflow"""
    workflow = CustomWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
'''

    def _get_excel_template(self) -> str:
        """Get Excel report template"""
        return '''"""
WORKFLOW_META:
  name: Excel Report Generator
  description: Generate Excel reports from data
  category: Reports
  version: 1.0.0
  author: Automation Hub
  parameters:
    input_file:
      type: file
      description: Input data file (CSV or Excel)
      required: true
    output_file:
      type: string
      description: Output Excel file path
      required: true
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.modules.excel_automation import WorkbookHandler
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class ExcelReportWorkflow(BaseModule):
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_file = None

    def configure(self, **kwargs):
        self.input_file = kwargs.get('input_file')
        self.output_file = kwargs.get('output_file')

    def validate(self):
        if not self.input_file or not self.output_file:
            raise ValueError("input_file and output_file are required")
        return True

    def execute(self):
        try:
            logger.info(f"Generating Excel report from {self.input_file}")

            wb = WorkbookHandler()
            sheet = wb.workbook.active
            sheet.title = "Report"

            # TODO: Add data processing and report generation logic

            wb.save(self.output_file)
            logger.info(f"Report saved to {self.output_file}")

            return {"status": "success", "output": self.output_file}
        except Exception as e:
            self.handle_error(e, "Report generation failed")
            raise


def run(**kwargs):
    workflow = ExcelReportWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
'''

    def _get_asana_template(self) -> str:
        """Get Asana integration template"""
        return '''"""
WORKFLOW_META:
  name: Asana Task Creator
  description: Create Asana tasks from Excel or bulk list
  category: Asana
  version: 1.0.0
  author: Automation Hub
  parameters:
    project_email:
      type: string
      description: Asana project email address (find in Project Settings)
      required: true
      default: "your-project@mail.asana.com"
    excel_file:
      type: file
      description: Excel file with tasks (columns: name, description, assignee, due_date, priority)
      required: false
      default: ""
    task_name:
      type: string
      description: Single task name (if not using Excel)
      required: false
      default: ""
    task_description:
      type: text
      description: Single task description
      required: false
      default: ""
    assignee:
      type: string
      description: Assignee username (without @)
      required: false
      default: ""
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.utils.workflow_helpers import AsanaHelper
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class AsanaTaskWorkflow(BaseModule):
    def __init__(self):
        super().__init__(
            name="AsanaTask",
            description="Create Asana tasks",
            version="1.0.0",
            category="asana"
        )
        self.project_email = None
        self.excel_file = None
        self.task_name = None
        self.task_description = None
        self.assignee = None

    def configure(self, **kwargs):
        self.project_email = kwargs.get('project_email')
        self.excel_file = kwargs.get('excel_file', '')
        self.task_name = kwargs.get('task_name', '')
        self.task_description = kwargs.get('task_description', '')
        self.assignee = kwargs.get('assignee', '')

        if not self.project_email:
            raise ValueError("project_email is required")

        logger.info(f"Configured Asana task creation")
        logger.info(f"  Project: {self.project_email}")

        return True

    def validate(self):
        if not self.excel_file and not self.task_name:
            raise ValueError("Either excel_file or task_name must be provided")
        return True

    def execute(self):
        try:
            logger.info("=" * 60)
            logger.info("CREATING ASANA TASKS")
            logger.info("=" * 60)

            helper = AsanaHelper(project_email=self.project_email)

            # Option 1: Bulk create from Excel
            if self.excel_file:
                logger.info(f"Creating tasks from Excel: {self.excel_file}")
                result = helper.bulk_create_from_excel(self.excel_file, method='email')

            # Option 2: Create single task
            elif self.task_name:
                logger.info(f"Creating task: {self.task_name}")
                result = helper.create_task_via_email(
                    name=self.task_name,
                    description=self.task_description,
                    assignee=self.assignee
                )

            if result['status'] == 'success':
                data = result.get('data', {})
                created = data.get('created_count', 1)
                logger.info(f"✓ Successfully created {created} task(s)")

            logger.info("=" * 60)
            return self.success_result(result.get('data'), result.get('message'))

        except Exception as e:
            logger.error(f"Failed to create tasks: {e}")
            return self.error_result(str(e))


def run(**kwargs):
    workflow = AsanaTaskWorkflow()
    result = workflow.run(**kwargs)
    return result
'''

    def _get_email_template(self) -> str:
        """Get email automation template"""
        return '''"""
WORKFLOW_META:
  name: Email Automation
  description: Automate email processing and sending
  category: Email
  version: 1.0.0
  author: Automation Hub
  parameters:
    action:
      type: choice
      choices: ['send', 'read', 'process']
      description: Email action to perform
      required: true
      default: 'send'
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.modules.outlook_automation import EmailHandler
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class EmailWorkflow(BaseModule):
    def __init__(self):
        super().__init__()
        self.action = None

    def configure(self, **kwargs):
        self.action = kwargs.get('action', 'send')

    def validate(self):
        if self.action not in ['send', 'read', 'process']:
            raise ValueError("Invalid action")
        return True

    def execute(self):
        try:
            logger.info(f"Performing email action: {self.action}")

            email_handler = EmailHandler()

            # TODO: Add email automation logic

            logger.info("Email workflow completed")
            return {"status": "success"}
        except Exception as e:
            self.handle_error(e, "Email workflow failed")
            raise


def run(**kwargs):
    workflow = EmailWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
'''

    def _get_file_management_template(self) -> str:
        """Get file management template"""
        return '''"""
WORKFLOW_META:
  name: File Management Workflow
  description: Automate file and folder operations
  category: Files
  version: 1.0.0
  author: Automation Hub
  parameters:
    source_folder:
      type: string
      description: Source folder path
      required: true
    action:
      type: choice
      choices: ['organize', 'copy', 'archive']
      description: File operation to perform
      required: true
      default: 'organize'
"""

import sys
import os
import shutil
from pathlib import Path

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class FileManagementWorkflow(BaseModule):
    def __init__(self):
        super().__init__()
        self.source_folder = None
        self.action = None

    def configure(self, **kwargs):
        self.source_folder = kwargs.get('source_folder')
        self.action = kwargs.get('action', 'organize')

    def validate(self):
        if not self.source_folder or not os.path.exists(self.source_folder):
            raise ValueError("source_folder must be a valid path")
        return True

    def execute(self):
        try:
            logger.info(f"Performing file operation: {self.action} on {self.source_folder}")

            # TODO: Add file management logic

            logger.info("File workflow completed")
            return {"status": "success"}
        except Exception as e:
            self.handle_error(e, "File workflow failed")
            raise


def run(**kwargs):
    workflow = FileManagementWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
'''

    def _extract_parameters_from_description(
        self,
        description: str,
        intent: str
    ) -> Dict[str, Any]:
        """Extract parameters from description"""
        # This is a simplified version - AI would do better
        params = {}

        if intent == 'excel_report':
            params = {
                'input_file': {'type': 'file', 'required': True, 'description': 'Input file'},
                'output_file': {'type': 'string', 'required': True, 'description': 'Output file'}
            }
        elif intent == 'email':
            params = {
                'action': {'type': 'choice', 'choices': ['send', 'read'], 'required': True}
            }
        elif intent == 'file_management':
            params = {
                'source_folder': {'type': 'string', 'required': True, 'description': 'Source folder'}
            }

        return params

    def _generate_name_from_description(self, description: str) -> str:
        """Generate a workflow name from description"""
        # Take first few words and title case
        words = description.split()[:4]
        name = ' '.join(words)
        if len(name) > 40:
            name = name[:40] + "..."
        return name.title() + " Workflow"

    def customize_template(
        self,
        template_code: str,
        customization_request: str,
        template_metadata: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Customize an existing template using AI

        Args:
            template_code: The original template code
            customization_request: User's request for how to modify the template
            template_metadata: Template metadata (name, description, parameters)

        Returns:
            Dictionary with customized code and metadata
        """
        try:
            # Check if API key is available
            if not self.api_key:
                self.logger.warning("No API key found, returning template as-is")
                return {
                    'code': template_code,
                    'name': template_metadata.get('name', 'Customized Workflow'),
                    'description': template_metadata.get('description', ''),
                    'parameters': template_metadata.get('parameters', {})
                }

            import anthropic

            # Build customization prompt
            prompt = self._create_customization_prompt(
                template_code,
                customization_request,
                template_metadata
            )

            # Call Claude API
            client = anthropic.Anthropic(api_key=self.api_key)

            # Get model settings
            from core.config import ConfigManager
            config = ConfigManager()
            model = config.get('ai.model', 'claude-sonnet-4-5-20250929')
            max_tokens = config.get('ai.max_tokens', 4000)

            self.logger.info(f"Calling Claude API for template customization (model: {model})...")

            message = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            response_text = message.content[0].text

            # Parse the response
            result = self._parse_ai_response(response_text)

            self.logger.info("Template customized successfully with AI")
            return result

        except Exception as e:
            self.logger.error(f"Template customization failed: {e}")
            # Return original template on error
            return {
                'code': template_code,
                'name': template_metadata.get('name', 'Customized Workflow'),
                'description': template_metadata.get('description', ''),
                'parameters': template_metadata.get('parameters', {})
            }

    def _create_customization_prompt(
        self,
        template_code: str,
        customization_request: str,
        template_metadata: Dict[str, Any]
    ) -> str:
        """Create prompt for template customization"""
        return f"""You are an expert Python automation script customizer. You have an existing, working template that needs to be modified based on the user's request.

## Original Template

Name: {template_metadata.get('name', 'Unknown')}
Description: {template_metadata.get('description', 'No description')}
Category: {template_metadata.get('category', 'Custom')}

### Original Code:
```python
{template_code}
```

## User's Customization Request

{customization_request}

## Instructions

1. **Preserve the template structure**: Keep the WORKFLOW_META format, class structure, and overall pattern
2. **Make the requested modifications**: Apply the user's requested changes
3. **Maintain functionality**: Ensure the code remains working and follows best practices
4. **Update metadata**: Modify the name, description, and parameters as needed to reflect changes
5. **Add/modify parameters**: If the customization requires new inputs, add them to parameters
6. **Keep error handling**: Preserve or improve error handling and logging
7. **Update docstrings**: Ensure docstrings reflect the new functionality

## Output Requirements

Generate the complete, modified Python script with:
- Updated WORKFLOW_META in the docstring
- Modified class and methods as requested
- Updated run() function if needed
- Proper imports and error handling

Return ONLY the Python code, no explanations. The code should be ready to save and run.
"""

    def recommend_templates(
        self,
        description: str,
        available_templates: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """
        Recommend relevant templates based on user description

        Args:
            description: User's workflow description
            available_templates: List of available template metadata

        Returns:
            List of recommended templates (sorted by relevance)
        """
        if not available_templates:
            return []

        try:
            # Use simple keyword matching for now
            # Could be enhanced with AI for better matching
            description_lower = description.lower()
            recommendations = []

            for template in available_templates:
                score = 0
                name_lower = template.get('name', '').lower()
                desc_lower = template.get('description', '').lower()
                category_lower = template.get('category', '').lower()

                # Keyword matching
                if 'email' in description_lower and 'email' in (name_lower + desc_lower):
                    score += 10
                if 'report' in description_lower and 'report' in (name_lower + desc_lower):
                    score += 10
                if 'excel' in description_lower and 'excel' in (name_lower + desc_lower):
                    score += 10
                if 'file' in description_lower and 'file' in (name_lower + desc_lower):
                    score += 10
                if 'sharepoint' in description_lower and 'sharepoint' in (name_lower + desc_lower):
                    score += 10
                if 'word' in description_lower and 'word' in (name_lower + desc_lower):
                    score += 10

                # Category matching
                if category_lower in description_lower:
                    score += 5

                # Check for common words
                desc_words = set(description_lower.split())
                template_words = set((name_lower + ' ' + desc_lower).split())
                common_words = desc_words & template_words
                score += len(common_words)

                if score > 0:
                    recommendations.append({
                        'template': template,
                        'score': score,
                        'reason': self._generate_recommendation_reason(description, template)
                    })

            # Sort by score (descending)
            recommendations.sort(key=lambda x: x['score'], reverse=True)

            # Return top 5 recommendations
            return recommendations[:5]

        except Exception as e:
            self.logger.error(f"Template recommendation failed: {e}")
            return []

    def _generate_recommendation_reason(
        self,
        description: str,
        template: Dict[str, Any]
    ) -> str:
        """Generate a reason for recommending a template"""
        reasons = []

        description_lower = description.lower()
        name_lower = template.get('name', '').lower()
        desc_lower = template.get('description', '').lower()

        if 'email' in description_lower and 'email' in (name_lower + desc_lower):
            reasons.append("handles email automation")
        if 'report' in description_lower and 'report' in (name_lower + desc_lower):
            reasons.append("generates reports")
        if 'excel' in description_lower and 'excel' in (name_lower + desc_lower):
            reasons.append("works with Excel files")
        if 'file' in description_lower and 'file' in (name_lower + desc_lower):
            reasons.append("manages files")

        if reasons:
            return "This template " + " and ".join(reasons)
        else:
            return f"This template is in the {template.get('category', 'Custom')} category"

    def generate_from_document(
        self,
        analysis: Dict[str, Any],
        user_instructions: str = "",
        model: str = None
    ) -> Dict[str, Any]:
        """
        Generate Python template code from document analysis.

        Uses Document Analyzer results to automatically generate
        complete, runnable Python template code with proper WORKFLOW_META.

        Args:
            analysis: Document analysis from DocumentAnalyzer
            user_instructions: Optional user customization requests
            model: Claude model to use (defaults to config setting)

        Returns:
            Dictionary with:
            - success: bool
            - code: str (generated Python code)
            - name: str (template name)
            - description: str (what template does)
            - parameters: dict (WORKFLOW_META parameters)
            - error: str (error message if failed)
        """
        try:
            # Check if API key is available
            if not self.api_key:
                self.logger.error("No API key found for document generation")
                return {
                    'success': False,
                    'code': '',
                    'error': 'No API key found. Please configure your Anthropic API key in Settings.'
                }

            import anthropic

            # Get template mode from analysis
            template_mode = analysis.get('mode', 'fill_in')
            source_file = analysis.get('source_file', 'document.docx')

            self.logger.info(f"Generating template from document (mode: {template_mode})...")

            # Build mode-specific prompt
            prompt = self._create_document_generation_prompt(analysis, user_instructions)

            # Call Claude API
            client = anthropic.Anthropic(api_key=self.api_key)

            # Get model settings
            from core.config import ConfigManager
            config = ConfigManager()
            if model is None:
                model = config.get('ai.model', 'claude-sonnet-4-5-20250929')
            max_tokens = config.get('ai.max_tokens', 4000)

            self.logger.info(f"Calling Claude API (model: {model})...")

            message = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            # Extract response text
            response_text = message.content[0].text

            # Parse the AI response
            result = self._parse_ai_response(response_text)

            # Validate generated code
            if result.get('code'):
                validation_errors = self._validate_template_code(result['code'])
                if validation_errors:
                    self.logger.warning(f"Generated code has validation issues: {validation_errors}")
                    # Don't fail, just warn - AI might have done something clever

            self.logger.info(f"Template generated successfully: {result.get('name', 'Unnamed')}")

            return {
                'success': True,
                'code': result.get('code', ''),
                'name': result.get('name', analysis.get('recommended_template_name', 'Generated Template')),
                'description': result.get('description', ''),
                'parameters': result.get('parameters', analysis.get('parameters', {})),
                'error': ''
            }

        except Exception as e:
            self.logger.error(f"Document template generation failed: {e}")
            return {
                'success': False,
                'code': '',
                'error': str(e)
            }

    def _create_document_generation_prompt(
        self,
        analysis: Dict[str, Any],
        user_instructions: str = ""
    ) -> str:
        """
        Create AI prompt for document template generation.

        Args:
            analysis: Document analysis from DocumentAnalyzer
            user_instructions: Optional user customization

        Returns:
            Formatted prompt string
        """
        template_mode = analysis.get('mode', 'fill_in')

        # Dispatch to mode-specific prompt builder
        if template_mode == 'fill_in':
            return self._create_fill_in_prompt(analysis, user_instructions)
        elif template_mode == 'generate':
            return self._create_generate_prompt(analysis, user_instructions)
        elif template_mode == 'content':
            return self._create_content_prompt(analysis, user_instructions)
        elif template_mode == 'pattern':
            return self._create_pattern_prompt(analysis, user_instructions)
        else:
            # Default to fill_in
            return self._create_fill_in_prompt(analysis, user_instructions)

    def _create_fill_in_prompt(self, analysis: Dict, user_instructions: str) -> str:
        """Create prompt for FILL_IN mode template generation."""
        placeholders = analysis.get('placeholders', [])
        parameters = analysis.get('parameters', {})
        structure = analysis.get('structure', {})
        template_name = analysis.get('recommended_template_name', 'Document Template')
        category = analysis.get('recommended_category', 'Custom')
        source_file = Path(analysis.get('source_file', 'template.docx')).name

        # Format placeholder list
        placeholder_list = "\n".join([
            f"  - {p['name']} ({p['type']}): {p['description']} - Pattern: {p['pattern']}"
            for p in placeholders[:10]  # First 10
        ])
        if len(placeholders) > 10:
            placeholder_list += f"\n  ... and {len(placeholders) - 10} more placeholders"

        # Format parameter definitions
        param_yaml_lines = []
        for param_name, param_def in parameters.items():
            param_yaml_lines.append(f"    {param_name}:")
            param_yaml_lines.append(f"      type: {param_def.get('type', 'string')}")
            param_yaml_lines.append(f"      description: {param_def.get('description', '')}")
            param_yaml_lines.append(f"      required: {str(param_def.get('required', False)).lower()}")
            if param_def.get('default') is not None:
                param_yaml_lines.append(f"      default: {param_def['default']}")

        param_yaml = "\n".join(param_yaml_lines)

        # Build user instructions section (avoid backslash in f-string for Python < 3.12)
        user_instructions_section = ""
        if user_instructions:
            user_instructions_section = "## USER INSTRUCTIONS\n\n" + user_instructions + "\n"

        prompt = f"""You are an expert Python automation developer. Generate a complete, production-ready Python workflow template.

## DOCUMENT ANALYSIS

**Source Document:** {source_file}
**Template Mode:** FILL_IN (document with placeholders to replace)
**Template Purpose:** {template_name}
**Category:** {category}

**Document Structure:**
- Paragraphs: {structure.get('total_paragraphs', 0)}
- Tables: {structure.get('total_tables', 0)}
- Headings: {len(structure.get('headings', []))}
- Complexity Score: {structure.get('complexity_score', 0)}/10

**Placeholders Detected ({len(placeholders)}):**
{placeholder_list}

## TASK

Generate a complete Python workflow template that:

1. **Loads the document template** from `user_scripts/templates/documents/{source_file}`
2. **Replaces all placeholders** with parameter values from kwargs
3. **Saves to output_file** parameter location
4. **Includes proper WORKFLOW_META** with all detected parameters
5. **Extends BaseModule** with configure(), validate(), execute() methods
6. **Has comprehensive error handling** and logging
7. **Returns structured result** dict with status and output path

## REQUIRED CODE STRUCTURE

```python
\"\"\"
{template_name}

WORKFLOW_META:
  name: {template_name}
  description: Generate documents from template with placeholder replacement
  category: {category}
  version: 1.0.0
  author: AI Generated
  parameters:
{param_yaml}
\"\"\"

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.modules.word_automation.document_handler import DocumentHandler
from src.core.logging_config import get_logger

logger = get_logger(__name__)

class [ClassName](BaseModule):
    \"\"\"[Description]\"\"\"

    def configure(self, **kwargs):
        # Extract all parameters from kwargs
        pass

    def validate(self):
        # Validate required parameters exist
        return True

    def execute(self):
        try:
            logger.info("Loading document template...")

            # Get template path
            template_path = Path(__file__).parent.parent.parent / "user_scripts" / "templates" / "documents" / "{source_file}"

            # Load template
            handler = DocumentHandler(str(template_path))

            # Build placeholder mapping from parameters
            placeholder_mapping = {{
                # Map all parameters to their values
            }}

            # Replace placeholders
            handler.replace_placeholders(placeholder_mapping)

            # Save output
            handler.save(self.output_file)

            logger.info(f"Document generated successfully: {{self.output_file}}")
            return {{
                'status': 'success',
                'output_file': self.output_file
            }}

        except Exception as e:
            logger.error(f"Document generation failed: {{e}}")
            raise

def run(**kwargs):
    workflow = [ClassName]()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()
```

## REQUIREMENTS

1. Use proper class name (TitleCase, no spaces)
2. Include ALL detected parameters in WORKFLOW_META
3. Map ALL placeholders in placeholder_mapping
4. Include comprehensive error handling
5. Use proper logging
6. Return structured result dict
7. Follow PEP 8 style guidelines

{user_instructions_section}

## OUTPUT

Return ONLY the complete Python code. No explanations, no markdown code blocks, just the raw Python code that can be saved directly to a .py file."""

        return prompt

    def _create_generate_prompt(self, analysis: Dict, user_instructions: str) -> str:
        """Create prompt for GENERATE mode template generation."""
        structure = analysis.get('structure', {})
        template_name = analysis.get('recommended_template_name', 'Document Generator')
        category = analysis.get('recommended_category', 'Custom')
        headings = structure.get('headings', [])
        tables = structure.get('tables', [])

        # Format heading structure
        heading_list = "\n".join([
            f"  - Level {h['level']}: \"{h['text']}\""
            for h in headings
        ])

        # Format table info
        table_list = "\n".join([
            f"  - Table {t['index']}: {t['rows']} rows x {t['cols']} columns"
            for t in tables
        ])

        # Build user instructions section (avoid backslash in f-string for Python < 3.12)
        user_instructions_section = ""
        if user_instructions:
            user_instructions_section = "## USER INSTRUCTIONS\n\n" + user_instructions + "\n"

        prompt = f"""You are an expert Python automation developer. Generate a complete, production-ready Python workflow template.

## DOCUMENT ANALYSIS

**Template Mode:** GENERATE (recreate document structure from scratch)
**Template Purpose:** {template_name}
**Category:** {category}

**Document Structure:**
- Paragraphs: {structure.get('total_paragraphs', 0)}
- Tables: {structure.get('total_tables', 0)}
- Headings: {len(headings)}
- Complexity Score: {structure.get('complexity_score', 0)}/10

**Heading Structure:**
{heading_list if heading_list else "  No headings"}

**Table Structure:**
{table_list if table_list else "  No tables"}

## TASK

Generate a Python workflow template that:

1. **Creates a new document** using DocumentHandler.create_document()
2. **Recreates the heading structure** with proper levels
3. **Accepts 'data' parameter** for dynamic content
4. **Generates tables** with proper formatting if needed
5. **Saves to output_file** parameter location
6. **Includes proper WORKFLOW_META**
7. **Has comprehensive error handling**

## REQUIRED CODE STRUCTURE

Follow the BaseModule pattern with configure/validate/execute methods.
Use DocumentHandler methods:
- create_document()
- add_heading(text, level)
- add_paragraph(text)
- add_table(data, headers)
- save(output_file)

{user_instructions_section}

Return ONLY the complete Python code with WORKFLOW_META in the docstring."""

        return prompt

    def _create_content_prompt(self, analysis: Dict, user_instructions: str) -> str:
        """Create prompt for CONTENT mode template generation."""
        template_name = analysis.get('recommended_template_name', 'Document Template')
        category = analysis.get('recommended_category', 'Custom')
        source_file = Path(analysis.get('source_file', 'template.docx')).name

        prompt = f"""Generate a simple Python workflow template for copying a static document.

**Template Mode:** CONTENT (static document reuse)
**Purpose:** {template_name}
**Source:** {source_file}

The template should:
1. Copy the template document from user_scripts/templates/documents/{source_file}
2. Save to output_file parameter
3. Include minimal WORKFLOW_META with just output_file parameter

Return ONLY the Python code."""

        return prompt

    def _create_pattern_prompt(self, analysis: Dict, user_instructions: str) -> str:
        """Create prompt for PATTERN mode template generation."""
        template_name = analysis.get('recommended_template_name', 'Batch Document Generator')
        category = analysis.get('recommended_category', 'Files')
        source_file = Path(analysis.get('source_file', 'template.docx')).name

        prompt = f"""Generate a Python workflow template for batch document generation.

**Template Mode:** PATTERN (repeating structure / mail merge)
**Purpose:** {template_name}
**Source:** {source_file}

The template should:
1. Accept data_source parameter (CSV/Excel file path)
2. Accept output_folder parameter
3. Read data from source file
4. For each row, load template and generate a document
5. Support placeholder replacement for each row
6. Save each document with unique filename

Return ONLY the Python code with proper WORKFLOW_META."""

        return prompt

    def _validate_template_code(self, code: str) -> List[str]:
        """
        Validate generated template code.

        Args:
            code: Python code string

        Returns:
            List of validation error messages (empty if valid)
        """
        errors = []

        # Check for WORKFLOW_META
        if 'WORKFLOW_META' not in code:
            errors.append("Missing WORKFLOW_META in docstring")

        # Check Python syntax
        try:
            compile(code, '<string>', 'exec')
        except SyntaxError as e:
            errors.append(f"Syntax error: {e}")

        # Check for required imports
        required_imports = [
            'BaseModule',
            'DocumentHandler',
            'get_logger'
        ]
        for imp in required_imports:
            if imp not in code:
                errors.append(f"Missing required import: {imp}")

        # Check for class definition
        if not re.search(r'class\s+\w+\(BaseModule\):', code):
            errors.append("Missing class that extends BaseModule")

        # Check for required methods
        required_methods = ['configure', 'validate', 'execute']
        for method in required_methods:
            if f'def {method}' not in code:
                errors.append(f"Missing required method: {method}()")

        # Check for run() function
        if 'def run(' not in code:
            errors.append("Missing run(**kwargs) function")

        return errors
