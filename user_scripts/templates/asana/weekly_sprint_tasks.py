"""
WORKFLOW_META:
  name: Create Weekly Sprint Tasks
  description: Automatically create weekly sprint tasks from templates
  category: Asana
  version: 1.0.0
  author: Automation Hub
  parameters:
    project_email:
      type: string
      description: Asana project email address
      required: true
      default: "sprint@mail.asana.com"
    sprint_number:
      type: integer
      description: Sprint number
      required: true
      default: 1
    start_date:
      type: string
      description: Sprint start date (YYYY-MM-DD)
      required: true
      default: "2024-12-09"
"""

import sys
from pathlib import Path
from datetime import datetime, timedelta

# Add src directory to path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger
from src.utils.workflow_helpers import AsanaHelper

logger = get_logger(__name__)


class WeeklySprintTasksWorkflow(BaseModule):
    """Create weekly sprint tasks automatically"""

    def __init__(self):
        super().__init__(
            name="WeeklySprintTasks",
            description="Create weekly sprint tasks",
            version="1.0.0",
            category="asana"
        )
        self.project_email = None
        self.sprint_number = None
        self.start_date = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.project_email = kwargs.get('project_email')
        self.sprint_number = kwargs.get('sprint_number')
        self.start_date = kwargs.get('start_date')

        if not self.project_email:
            raise ValueError("project_email is required")
        if not self.sprint_number:
            raise ValueError("sprint_number is required")
        if not self.start_date:
            raise ValueError("start_date is required")

        logger.info(f"Configured sprint task creation:")
        logger.info(f"  Sprint: {self.sprint_number}")
        logger.info(f"  Start date: {self.start_date}")
        logger.info(f"  Project email: {self.project_email}")

        return True

    def validate(self):
        """Validate configuration"""
        # Validate date format
        try:
            datetime.strptime(self.start_date, '%Y-%m-%d')
        except ValueError:
            raise ValueError(f"Invalid date format: {self.start_date}. Use YYYY-MM-DD")

        return True

    def execute(self):
        """Execute the workflow"""
        logger.info("=" * 60)
        logger.info(f"CREATING SPRINT {self.sprint_number} TASKS")
        logger.info("=" * 60)

        # Define sprint task templates
        task_templates = [
            {
                'name': 'Sprint Planning',
                'assignee': 'team',
                'day_offset': 0,
                'description': 'Plan sprint goals and select backlog items',
                'priority': 'High'
            },
            {
                'name': 'Daily Standup - Monday',
                'assignee': 'team',
                'day_offset': 0,
                'description': '15-minute sync: What did you do? What will you do? Any blockers?'
            },
            {
                'name': 'Daily Standup - Tuesday',
                'assignee': 'team',
                'day_offset': 1
            },
            {
                'name': 'Daily Standup - Wednesday',
                'assignee': 'team',
                'day_offset': 2
            },
            {
                'name': 'Daily Standup - Thursday',
                'assignee': 'team',
                'day_offset': 3
            },
            {
                'name': 'Daily Standup - Friday',
                'assignee': 'team',
                'day_offset': 4
            },
            {
                'name': 'Sprint Review',
                'assignee': 'team',
                'day_offset': 4,
                'description': 'Demo completed work to stakeholders',
                'priority': 'High'
            },
            {
                'name': 'Sprint Retrospective',
                'assignee': 'team',
                'day_offset': 4,
                'description': 'Reflect on sprint: What went well? What can improve?',
                'priority': 'High'
            },
            {
                'name': 'Update Sprint Metrics',
                'assignee': 'scrum-master',
                'day_offset': 4,
                'description': 'Update velocity, burndown, and other sprint metrics'
            }
        ]

        # Initialize helper
        helper = AsanaHelper(project_email=self.project_email)

        # Create sprint tasks
        logger.info(f"Creating {len(task_templates)} sprint tasks...")
        result = helper.create_weekly_sprint_tasks(
            sprint_number=self.sprint_number,
            start_date=self.start_date,
            task_templates=task_templates
        )

        # Log results
        if result['status'] == 'success':
            data = result['data']
            logger.info(f"✓ Successfully created {data['created_count']} tasks")

            if data.get('failed_count', 0) > 0:
                logger.warning(f"✗ Failed to create {data['failed_count']} tasks")

        logger.info("=" * 60)
        logger.info("SPRINT TASKS CREATED")
        logger.info("=" * 60)

        return result


def main():
    """Main entry point"""
    workflow = WeeklySprintTasksWorkflow()

    # Get current week's Monday
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())

    # Example configuration
    result = workflow.run(
        project_email='sprint@mail.asana.com',
        sprint_number=42,
        start_date=monday.strftime('%Y-%m-%d')
    )

    print(f"\nResult: {result['message']}")


if __name__ == '__main__':
    main()
