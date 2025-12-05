"""
File Helpers

Helper classes and functions for file and document management.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Callable, Any

from core.logging_config import get_logger

logger = get_logger(__name__)


class FileOrganizer:
    """Helper for file management tasks."""

    def __init__(self) -> None:
        pass

    def organize_by_date(
        self,
        source_folder: str,
        dest_folder: str,
        date_format: str = "%Y-%m"
    ) -> None:
        """
        Organize files into folders by date.

        Args:
            source_folder: Source directory.
            dest_folder: Destination directory.
            date_format: Date format for folder names (e.g., "2024-01").
        """
        os.makedirs(dest_folder, exist_ok=True)

        files = Path(source_folder).glob('*')
        moved_count = 0

        for file_path in files:
            if file_path.is_file():
                # Get file date
                file_date = datetime.fromtimestamp(file_path.stat().st_mtime)
                date_folder = file_date.strftime(date_format)

                # Create date folder
                target_folder = os.path.join(dest_folder, date_folder)
                os.makedirs(target_folder, exist_ok=True)

                # Move file
                target_path = os.path.join(target_folder, file_path.name)
                shutil.move(str(file_path), target_path)
                logger.info(f"Moved: {file_path.name} -> {date_folder}/")
                moved_count += 1

        logger.info(f"Organized {moved_count} files by date")

    def organize_by_type(
        self,
        source_folder: str,
        dest_folder: str,
        type_mapping: Optional[Dict[str, List[str]]] = None
    ) -> None:
        """
        Organize files by type/extension.

        Args:
            source_folder: Source directory.
            dest_folder: Destination directory.
            type_mapping: Dict of {folder_name: [extensions]}.
                         e.g., {"Documents": [".pdf", ".docx"], "Images": [".jpg", ".png"]}
        """
        if type_mapping is None:
            type_mapping = {
                "Documents": [".pdf", ".doc", ".docx", ".txt"],
                "Spreadsheets": [".xlsx", ".xls", ".csv"],
                "Images": [".jpg", ".jpeg", ".png", ".gif"],
                "Archives": [".zip", ".rar", ".7z"],
                "Other": []  # Catch-all
            }

        os.makedirs(dest_folder, exist_ok=True)

        files = Path(source_folder).glob('*')
        moved_count = 0

        for file_path in files:
            if file_path.is_file():
                ext = file_path.suffix.lower()

                # Find matching type
                target_folder_name = "Other"
                for folder_name, extensions in type_mapping.items():
                    if ext in extensions:
                        target_folder_name = folder_name
                        break

                # Create type folder
                target_folder = os.path.join(dest_folder, target_folder_name)
                os.makedirs(target_folder, exist_ok=True)

                # Move file
                target_path = os.path.join(target_folder, file_path.name)
                shutil.move(str(file_path), target_path)
                logger.info(f"Moved: {file_path.name} -> {target_folder_name}/")
                moved_count += 1

        logger.info(f"Organized {moved_count} files by type")

    def batch_rename(
        self,
        folder: str,
        pattern: str,
        replacement: str
    ) -> None:
        """
        Batch rename files in a folder.

        Args:
            folder: Folder containing files.
            pattern: Pattern to replace.
            replacement: Replacement string.
        """
        files = Path(folder).glob('*')
        renamed_count = 0

        for file_path in files:
            if file_path.is_file() and pattern in file_path.name:
                new_name = file_path.name.replace(pattern, replacement)
                new_path = file_path.parent / new_name

                file_path.rename(new_path)
                logger.info(f"Renamed: {file_path.name} -> {new_name}")
                renamed_count += 1

        logger.info(f"Renamed {renamed_count} files")

    def archive_old_files(
        self,
        source_folder: str,
        archive_folder: str,
        days_old: int = 90
    ) -> None:
        """
        Archive files older than specified days.

        Args:
            source_folder: Source directory.
            archive_folder: Archive directory.
            days_old: Archive files older than this many days.
        """
        os.makedirs(archive_folder, exist_ok=True)

        cutoff_date = datetime.now().timestamp() - (days_old * 24 * 60 * 60)
        files = Path(source_folder).glob('*')
        archived_count = 0

        for file_path in files:
            if file_path.is_file():
                file_date = file_path.stat().st_mtime

                if file_date < cutoff_date:
                    target_path = os.path.join(archive_folder, file_path.name)
                    shutil.move(str(file_path), target_path)
                    logger.info(f"Archived: {file_path.name}")
                    archived_count += 1

        logger.info(f"Archived {archived_count} old files")


class SharePointHelper:
    """Helper for SharePoint operations."""

    def __init__(self, site_url: str) -> None:
        from modules.sharepoint import SharePointClient
        self.client = SharePointClient(site_url)

    def sync_folder_to_sharepoint(
        self,
        local_folder: str,
        sp_folder: str
    ) -> None:
        """
        Sync a local folder to SharePoint.

        Args:
            local_folder: Local folder path.
            sp_folder: SharePoint folder path.
        """
        files = Path(local_folder).glob('*')
        uploaded_count = 0

        for file_path in files:
            if file_path.is_file():
                sp_path = f"{sp_folder}/{file_path.name}"
                self.client.upload_file(str(file_path), sp_path)
                logger.info(f"Uploaded: {file_path.name}")
                uploaded_count += 1

        logger.info(f"Synced {uploaded_count} files to SharePoint")

    def download_recent_files(
        self,
        sp_folder: str,
        local_folder: str,
        days: int = 7
    ) -> None:
        """
        Download files from SharePoint modified in the last N days.

        Args:
            sp_folder: SharePoint folder path.
            local_folder: Local destination folder.
            days: Download files modified in the last N days.
        """
        os.makedirs(local_folder, exist_ok=True)

        # Get files from SharePoint
        files = self.client.list_files(sp_folder)
        cutoff_date = datetime.now().timestamp() - (days * 24 * 60 * 60)

        downloaded_count = 0
        for file_info in files:
            if file_info.get('modified', 0) > cutoff_date:
                sp_path = file_info['path']
                local_path = os.path.join(local_folder, file_info['name'])

                self.client.download_file(sp_path, local_path)
                logger.info(f"Downloaded: {file_info['name']}")
                downloaded_count += 1

        logger.info(f"Downloaded {downloaded_count} recent files")


def run_file_processing_workflow(
    input_folder: str,
    output_folder: str,
    processing_func: Callable[[str], Any],
    organize_by_date: bool = False
) -> None:
    """
    Common pattern: Process files in a folder.

    Args:
        input_folder: Input folder path.
        output_folder: Output folder path.
        processing_func: Function to process each file.
        organize_by_date: Whether to organize outputs by date.
    """
    logger.info("Running file processing workflow...")

    os.makedirs(output_folder, exist_ok=True)
    files = Path(input_folder).glob('*')

    processed_count = 0
    for file_path in files:
        if file_path.is_file():
            try:
                # Process file
                result = processing_func(str(file_path))

                # Save result
                output_path = os.path.join(output_folder, file_path.name)
                # Save result logic here

                logger.info(f"Processed: {file_path.name}")
                processed_count += 1

            except Exception as e:
                logger.error(f"Failed to process {file_path.name}: {e}")

    if organize_by_date:
        organizer = FileOrganizer()
        organizer.organize_by_date(output_folder, output_folder)

    logger.info(f"File processing workflow completed! Processed {processed_count} files")
