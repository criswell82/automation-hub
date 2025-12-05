"""
WORKFLOW_META:
  name: File Organizer and SharePoint Sync
  description: Organize files by type/date and optionally sync to SharePoint
  category: Files
  version: 1.0.0
  author: Automation Hub
  parameters:
    source_folder:
      type: string
      description: Source folder containing files to organize
      required: true
      default: "C:/Downloads"
    destination_folder:
      type: string
      description: Destination folder for organized files
      required: true
      default: "C:/OrganizedFiles"
    organize_method:
      type: choice
      choices: ['by_type', 'by_date', 'both']
      description: How to organize files
      required: true
      default: 'by_type'
    archive_old_files:
      type: boolean
      description: Archive files older than 90 days
      required: false
      default: false
    sync_to_sharepoint:
      type: boolean
      description: Sync organized files to SharePoint
      required: false
      default: false
    sharepoint_site:
      type: string
      description: SharePoint site URL (if syncing)
      required: false
      default: ""
    sharepoint_folder:
      type: string
      description: SharePoint folder path (if syncing)
      required: false
      default: "/Shared Documents"
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
from src.utils.workflow_helpers import FileOrganizer, SharePointHelper

logger = get_logger(__name__)


class FileOrganizerWorkflow(BaseModule):
    """Organize files and optionally sync to SharePoint"""

    def __init__(self):
        super().__init__()
        self.source_folder = None
        self.destination_folder = None
        self.organize_method = None
        self.archive_old_files = None
        self.sync_to_sharepoint = None
        self.sharepoint_site = None
        self.sharepoint_folder = None

    def configure(self, **kwargs):
        """Configure the workflow"""
        self.source_folder = kwargs.get('source_folder')
        self.destination_folder = kwargs.get('destination_folder')
        self.organize_method = kwargs.get('organize_method', 'by_type')
        self.archive_old_files = kwargs.get('archive_old_files', False)
        self.sync_to_sharepoint = kwargs.get('sync_to_sharepoint', False)
        self.sharepoint_site = kwargs.get('sharepoint_site', '')
        self.sharepoint_folder = kwargs.get('sharepoint_folder', '/Shared Documents')

        logger.info(f"Configured file organizer workflow:")
        logger.info(f"  Source: {self.source_folder}")
        logger.info(f"  Destination: {self.destination_folder}")
        logger.info(f"  Method: {self.organize_method}")

    def validate(self):
        """Validate configuration"""
        if not self.source_folder or not os.path.exists(self.source_folder):
            raise ValueError(f"source_folder must exist: {self.source_folder}")

        if not self.destination_folder:
            raise ValueError("destination_folder is required")

        if self.sync_to_sharepoint and not self.sharepoint_site:
            raise ValueError("sharepoint_site is required when sync_to_sharepoint is true")

        return True

    def execute(self):
        """Main workflow logic"""
        try:
            logger.info("Starting file organization workflow...")

            # Initialize file organizer
            organizer = FileOrganizer()

            # Count initial files
            initial_files = len(list(Path(self.source_folder).glob('*')))
            logger.info(f"Found {initial_files} files to organize")

            # Archive old files first if requested
            if self.archive_old_files:
                archive_folder = os.path.join(self.destination_folder, "_Archive")
                organizer.archive_old_files(
                    source_folder=self.source_folder,
                    archive_folder=archive_folder,
                    days_old=90
                )

            # Organize files
            if self.organize_method == 'by_type':
                organizer.organize_by_type(
                    source_folder=self.source_folder,
                    dest_folder=self.destination_folder
                )

            elif self.organize_method == 'by_date':
                organizer.organize_by_date(
                    source_folder=self.source_folder,
                    dest_folder=self.destination_folder,
                    date_format="%Y-%m"
                )

            elif self.organize_method == 'both':
                # First by type
                temp_folder = os.path.join(self.destination_folder, "_temp")
                organizer.organize_by_type(
                    source_folder=self.source_folder,
                    dest_folder=temp_folder
                )

                # Then by date within each type
                type_folders = [f for f in Path(temp_folder).iterdir() if f.is_dir()]
                for type_folder in type_folders:
                    final_folder = os.path.join(self.destination_folder, type_folder.name)
                    organizer.organize_by_date(
                        source_folder=str(type_folder),
                        dest_folder=final_folder,
                        date_format="%Y-%m"
                    )

                # Clean up temp folder
                import shutil
                shutil.rmtree(temp_folder)

            logger.info("File organization completed!")

            # Sync to SharePoint if requested
            sp_synced = 0
            if self.sync_to_sharepoint:
                logger.info("Syncing to SharePoint...")
                try:
                    sp_helper = SharePointHelper(self.sharepoint_site)
                    sp_helper.sync_folder_to_sharepoint(
                        local_folder=self.destination_folder,
                        sp_folder=self.sharepoint_folder
                    )
                    sp_synced = len(list(Path(self.destination_folder).glob('**/*')))
                    logger.info("SharePoint sync completed!")
                except Exception as e:
                    logger.error(f"SharePoint sync failed: {e}")

            # Count final organized files
            organized_files = len(list(Path(self.destination_folder).glob('**/*')))

            logger.info("File organization workflow completed!")

            return {
                "status": "success",
                "message": "Files organized successfully",
                "initial_files": initial_files,
                "organized_files": organized_files,
                "destination": self.destination_folder,
                "sharepoint_synced": sp_synced if self.sync_to_sharepoint else 0
            }

        except Exception as e:
            self.handle_error(e, "File organization failed")
            raise


def run(**kwargs):
    """Execute the workflow"""
    workflow = FileOrganizerWorkflow()
    workflow.configure(**kwargs)
    workflow.validate()
    return workflow.execute()


if __name__ == "__main__":
    # Test the workflow
    result = run(
        source_folder="C:/Downloads",
        destination_folder="C:/OrganizedFiles",
        organize_method='by_type',
        archive_old_files=False,
        sync_to_sharepoint=False
    )
    print(f"\nResult: {result}")
