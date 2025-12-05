"""
WORKFLOW_META:
  name: Hello World Example
  description: A simple example custom workflow to test the system
  category: Custom
  version: 1.0.0
  author: Automation Hub
  parameters:
    name:
      type: string
      description: Your name
      required: false
      default: "User"
    message:
      type: text
      description: Custom message to display
      required: false
      default: "Welcome to Automation Hub!"
"""

import sys
from pathlib import Path

# Add src directory to path for imports
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class HelloWorldWorkflow(BaseModule):
    """Example custom workflow"""

    def __init__(self):
        super().__init__()
        self.name = None
        self.message = None

    def configure(self, **kwargs):
        """Configure the workflow with parameters"""
        self.name = kwargs.get('name', 'User')
        self.message = kwargs.get('message', 'Welcome to Automation Hub!')
        logger.info(f"Configured HelloWorld workflow for {self.name}")

    def validate(self):
        """Validate configuration"""
        if not self.name:
            raise ValueError("name parameter is required")
        return True

    def execute(self):
        """Main workflow logic"""
        try:
            logger.info("Starting Hello World workflow...")

            # Print greeting
            greeting = f"Hello, {self.name}!"
            print("=" * 50)
            print(greeting)
            print("-" * 50)
            print(self.message)
            print("=" * 50)

            logger.info("Workflow completed successfully!")

            return {
                "status": "success",
                "message": f"Greeted {self.name} successfully",
                "greeting": greeting
            }

        except Exception as e:
            self.handle_error(e, "Hello World workflow failed")
            raise


def run(**kwargs):
    """Execute the workflow"""
    workflow = HelloWorldWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()


if __name__ == "__main__":
    # Test the workflow
    result = run(name="Test User", message="This is a test!")
    print(f"\nResult: {result}")
