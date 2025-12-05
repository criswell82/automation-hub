"""
WORKFLOW_META:
  name: Generate Asana Import CSV
  description: Convert Excel or custom CSV to Asana-compatible CSV for manual import
  category: Asana
  version: 1.0.0
  author: Automation Hub
  parameters:
    input_file:
      type: string
      description: Path to input file (Excel or CSV)
      required: true
      default: "C:/Data/project_tasks.xlsx"
    output_csv:
      type: string
      description: Path for output Asana CSV
      required: true
      default: "C:/Output/asana_import.csv"
    column_mapping:
      type: string
      description: Optional column mapping (JSON format)
      required: false
      default: ""
"""

import sys
import json
from pathlib import Path

# Add src directory to path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger
from src.utils.workflow_helpers import AsanaHelper

logger = get_logger(__name__)


class GenerateAsanaCSVWorkflow(BaseModule):
    """Generate Asana-compatible CSV from Excel or custom CSV"""

    def __init__(self):
        super().__init__(
            name="GenerateAsanaCSV",
            description="Generate Asana import CSV",
            version="1.0.0",
            category="asana"
        )
        self.input_file = None
        self.output_csv = None
        self.column_mapping = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.input_file = kwargs.get('input_file')
        self.output_csv = kwargs.get('output_csv')

        mapping_str = kwargs.get('column_mapping', '')
        if mapping_str:
            try:
                self.column_mapping = json.loads(mapping_str)
            except json.JSONDecodeError:
                logger.warning(f"Invalid JSON for column_mapping, using auto-detection")
                self.column_mapping = {}
        else:
            self.column_mapping = {}

        if not self.input_file:
            raise ValueError("input_file is required")
        if not self.output_csv:
            raise ValueError("output_csv is required")

        logger.info(f"Configured CSV generation:")
        logger.info(f"  Input file: {self.input_file}")
        logger.info(f"  Output CSV: {self.output_csv}")
        if self.column_mapping:
            logger.info(f"  Column mapping: {self.column_mapping}")

        return True

    def validate(self):
        """Validate configuration"""
        # Check input file exists
        if not Path(self.input_file).exists():
            raise FileNotFoundError(f"Input file not found: {self.input_file}")

        # Check file type
        ext = Path(self.input_file).suffix.lower()
        if ext not in ['.xlsx', '.xls', '.csv']:
            raise ValueError(f"Unsupported file type: {ext}. Use .xlsx, .xls, or .csv")

        return True

    def execute(self):
        """Execute the workflow"""
        logger.info("=" * 60)
        logger.info("GENERATE ASANA IMPORT CSV")
        logger.info("=" * 60)

        # Initialize helper
        helper = AsanaHelper()

        # Convert to Asana CSV
        logger.info(f"Converting {self.input_file} to Asana format...")

        if self.column_mapping:
            result = helper.convert_excel_to_asana_csv(
                excel_path=self.input_file,
                output_path=self.output_csv,
                column_mapping=self.column_mapping
            )
        else:
            # Auto-detection
            from src.modules.asana import AsanaCSVHandler

            handler = AsanaCSVHandler()
            result = handler.run(
                operation='generate',
                input_file=self.input_file,
                output_file=self.output_csv
            )

        # Log results
        if result['status'] == 'success':
            data = result['data']
            logger.info(f"✓ Generated Asana CSV with {data['task_count']} tasks")
            logger.info(f"✓ Output: {data['output_file']}")
            logger.info("")
            logger.info("NEXT STEPS:")
            logger.info("1. Open Asana project in web browser")
            logger.info("2. Click project dropdown > Import > CSV")
            logger.info(f"3. Select file: {data['output_file']}")
            logger.info("4. Map columns and import")

        logger.info("=" * 60)
        logger.info("CSV GENERATION COMPLETE")
        logger.info("=" * 60)

        return result


def main():
    """Main entry point"""
    workflow = GenerateAsanaCSVWorkflow()

    # Example: Custom column mapping
    mapping = {
        'Task Name': 'title',
        'Assignee': 'owner',
        'Due Date': 'deadline',
        'Notes': 'description'
    }

    # Example configuration
    result = workflow.run(
        input_file='C:/Data/project_tasks.xlsx',
        output_csv='C:/Output/asana_import.csv',
        column_mapping=json.dumps(mapping)  # Convert dict to JSON string
    )

    print(f"\nResult: {result['message']}")
    print(f"CSV file: {result['data']['output_file']}")


if __name__ == '__main__':
    main()
