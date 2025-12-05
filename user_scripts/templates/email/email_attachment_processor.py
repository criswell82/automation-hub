"""
WORKFLOW_META:
  name: Email Attachment Processor
  description: Process emails with attachments - extract, organize, and optionally upload to SharePoint
  category: Email
  version: 1.0.0
  author: Automation Hub
  parameters:
    email_folder:
      type: string
      description: Outlook folder to process
      required: false
      default: "Inbox"
    subject_filter:
      type: string
      description: Filter emails by subject keyword (optional)
      required: false
      default: ""
    output_folder:
      type: string
      description: Folder to save attachments
      required: true
      default: "C:/EmailAttachments"
    file_types:
      type: string
      description: File extensions to extract (comma-separated, e.g., .pdf,.xlsx)
      required: false
      default: ".pdf,.xlsx,.docx"
    unread_only:
      type: boolean
      description: Only process unread emails
      required: false
      default: true
    organize_by_date:
      type: boolean
      description: Organize attachments into date folders
      required: false
      default: true
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# Add src directory to path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger
from src.utils.workflow_helpers import EmailHelper, FileOrganizer

logger = get_logger(__name__)


class EmailAttachmentProcessorWorkflow(BaseModule):
    """Process email attachments and organize them"""

    def __init__(self):
        super().__init__()
        self.email_folder = None
        self.subject_filter = None
        self.output_folder = None
        self.file_types = None
        self.unread_only = None
        self.organize_by_date = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.email_folder = kwargs.get('email_folder', 'Inbox')
        self.subject_filter = kwargs.get('subject_filter', '')
        self.output_folder = kwargs.get('output_folder')
        self.file_types = kwargs.get('file_types', '.pdf,.xlsx,.docx')
        self.unread_only = kwargs.get('unread_only', True)
        self.organize_by_date = kwargs.get('organize_by_date', True)

        logger.info(f"Configured email attachment processor:")
        logger.info(f"  Email folder: {self.email_folder}")
        logger.info(f"  Subject filter: {self.subject_filter or 'None'}")
        logger.info(f"  Output folder: {self.output_folder}")
        logger.info(f"  File types: {self.file_types}")

    def validate(self):
        """Validate configuration"""
        if not self.output_folder:
            raise ValueError("output_folder is required")

        return True

    def execute(self):
        """Main workflow logic"""
        try:
            logger.info("Starting email attachment processing...")

            # Initialize email helper
            email_helper = EmailHelper()

            # Process inbox
            emails = email_helper.process_inbox(
                folder=self.email_folder,
                unread_only=self.unread_only,
                filter_subject=self.subject_filter
            )

            logger.info(f"Found {len(emails)} emails to process")

            # Extract attachments
            file_extensions = [ext.strip() for ext in self.file_types.split(',') if ext.strip()]

            saved_files = email_helper.extract_attachments(
                emails=emails,
                output_folder=self.output_folder,
                file_extensions=file_extensions
            )

            # Organize files if requested
            if self.organize_by_date and saved_files:
                logger.info("Organizing attachments by date...")
                organizer = FileOrganizer()
                organizer.organize_by_date(
                    source_folder=self.output_folder,
                    dest_folder=self.output_folder,
                    date_format="%Y-%m"
                )

            logger.info("Email attachment processing completed!")

            return {
                "status": "success",
                "message": "Email attachments processed successfully",
                "emails_processed": len(emails),
                "attachments_saved": len(saved_files),
                "output_folder": self.output_folder
            }

        except Exception as e:
            self.handle_error(e, "Email attachment processing failed")
            raise


def run(**kwargs):
    """Execute the workflow"""
    workflow = EmailAttachmentProcessorWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()


if __name__ == "__main__":
    # Test the workflow
    result = run(
        email_folder="Inbox",
        subject_filter="Invoice",
        output_folder="C:/EmailAttachments",
        file_types=".pdf,.xlsx",
        unread_only=True,
        organize_by_date=True
    )
    print(f"\nResult: {result}")
