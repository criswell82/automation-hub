"""
WORKFLOW_META:
  name: Excel to Asana Bulk Import
  description: Bulk create Asana tasks from an Excel file via email integration
  category: Asana
  version: 1.0.0
  author: Automation Hub
  parameters:
    project_email:
      type: string
      description: Asana project email address (found in Project Settings)
      required: true
      default: "your-project@mail.asana.com"
    excel_file:
      type: string
      description: Path to Excel file with tasks (columns: name, description, assignee, due_date, priority)
      required: true
      default: "C:/Data/tasks.xlsx"
"""

import sys
from pathlib import Path
from datetime import datetime

# Add src directory to path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger
from src.utils.workflow_helpers import AsanaHelper

logger = get_logger(__name__)


class ExcelToAsanaWorkflow(BaseModule):
    """Bulk create Asana tasks from Excel file"""

    def __init__(self):
        super().__init__(
            name="ExcelToAsana",
            description="Bulk create Asana tasks from Excel",
            version="1.0.0",
            category="asana"
        )
        self.project_email = None
        self.excel_file = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.project_email = kwargs.get('project_email')
        self.excel_file = kwargs.get('excel_file')

        if not self.project_email:
            raise ValueError("project_email is required")
        if not self.excel_file:
            raise ValueError("excel_file is required")

        logger.info(f"Configured Excel to Asana import:")
        logger.info(f"  Project email: {self.project_email}")
        logger.info(f"  Excel file: {self.excel_file}")

        return True

    def validate(self):
        """Validate configuration"""
        # Check file exists
        if not Path(self.excel_file).exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_file}")

        return True

    def execute(self):
        """Execute the workflow"""
        logger.info("=" * 60)
        logger.info("EXCEL TO ASANA BULK IMPORT")
        logger.info("=" * 60)

        # Initialize helper
        helper = AsanaHelper(project_email=self.project_email)

        # Bulk create tasks from Excel
        logger.info(f"Creating tasks from Excel file: {self.excel_file}")
        result = helper.bulk_create_from_excel(self.excel_file, method='email')

        # Log results
        if result['status'] == 'success':
            data = result['data']
            logger.info(f"✓ Successfully created {data['created_count']} tasks")

            if data['failed_count'] > 0:
                logger.warning(f"✗ Failed to create {data['failed_count']} tasks")
                for failed in data['failed_tasks']:
                    logger.warning(f"  - {failed['name']}: {failed['error']}")

        logger.info("=" * 60)
        logger.info("WORKFLOW COMPLETE")
        logger.info("=" * 60)

        return result


def main():
    """Main entry point"""
    workflow = ExcelToAsanaWorkflow()

    # Example configuration
    result = workflow.run(
        project_email='your-project@mail.asana.com',
        excel_file='C:/Data/tasks.xlsx'
    )

    print(f"\nResult: {result['message']}")


if __name__ == '__main__':
    main()
