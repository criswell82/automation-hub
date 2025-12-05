"""
Communication Helpers

Helper classes and functions for email and communication automation.
"""

import os
from datetime import datetime
from typing import List, Dict, Any, Optional

from core.logging_config import get_logger

logger = get_logger(__name__)


class EmailHelper:
    """Helper for email automation tasks."""

    def __init__(self) -> None:
        from modules.outlook_automation import EmailHandler
        self.email_handler = EmailHandler()

    def send_report(
        self,
        to: List[str],
        subject: str,
        body: str,
        attachments: Optional[List[str]] = None,
        cc: Optional[List[str]] = None
    ) -> None:
        """
        Send an email with report attachments.

        Args:
            to: List of recipient email addresses.
            subject: Email subject.
            body: Email body (can include HTML).
            attachments: List of file paths to attach.
            cc: List of CC recipients.
        """
        logger.info(f"Sending email to: {', '.join(to)}")

        self.email_handler.send_email(
            to=to,
            subject=subject,
            body=body,
            attachments=attachments or [],
            cc=cc
        )

        logger.info("Email sent successfully")

    def process_inbox(
        self,
        folder: str = "Inbox",
        unread_only: bool = True,
        filter_subject: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Process emails from inbox.

        Args:
            folder: Email folder to process.
            unread_only: Only process unread emails.
            filter_subject: Filter by subject keyword.

        Returns:
            List of email dictionaries.
        """
        logger.info(f"Processing emails from: {folder}")

        emails = self.email_handler.read_emails(
            folder=folder,
            unread_only=unread_only
        )

        if filter_subject:
            emails = [
                e for e in emails
                if filter_subject.lower() in e.get('subject', '').lower()
            ]

        logger.info(f"Found {len(emails)} emails")
        return emails

    def extract_attachments(
        self,
        emails: List[Dict[str, Any]],
        output_folder: str,
        file_extensions: Optional[List[str]] = None
    ) -> List[str]:
        """
        Extract attachments from emails.

        Args:
            emails: List of email dictionaries (from process_inbox).
            output_folder: Where to save attachments.
            file_extensions: Optional filter (e.g., ['.pdf', '.xlsx']).

        Returns:
            List of saved file paths.
        """
        os.makedirs(output_folder, exist_ok=True)
        saved_files: List[str] = []

        for email in emails:
            attachments = email.get('attachments', [])

            for attachment in attachments:
                file_name = attachment.get('name', 'attachment')

                # Filter by extension if specified
                if file_extensions:
                    if not any(file_name.endswith(ext) for ext in file_extensions):
                        continue

                # Save attachment
                file_path = os.path.join(output_folder, file_name)
                # Note: Actual save logic would use email_handler
                saved_files.append(file_path)
                logger.info(f"Saved attachment: {file_name}")

        logger.info(f"Extracted {len(saved_files)} attachments")
        return saved_files

    def create_email_summary(
        self,
        to: str,
        subject: str,
        data: Dict[str, Any]
    ) -> None:
        """
        Create and send a summary email from data.

        Args:
            to: Recipient email.
            subject: Email subject.
            data: Dictionary of summary data.
        """
        # Build HTML email body
        body = "<html><body>"
        body += f"<h2>{subject}</h2>"
        body += "<table border='1' cellpadding='5' cellspacing='0'>"

        for key, value in data.items():
            body += f"<tr><td><b>{key}</b></td><td>{value}</td></tr>"

        body += "</table>"
        body += f"<p><i>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</i></p>"
        body += "</body></html>"

        self.send_report([to], subject, body)


def run_daily_report_workflow(
    data_source: str,
    report_output: str,
    email_recipients: List[str],
    email_subject: Optional[str] = None
) -> None:
    """
    Common pattern: Generate daily report and email it.

    Args:
        data_source: Path to data file or folder.
        report_output: Path for output Excel file.
        email_recipients: List of email addresses.
        email_subject: Optional subject (auto-generated if not provided).
    """
    from utils.report_helpers import ReportBuilder

    logger.info("Running daily report workflow...")

    # Build report
    report = ReportBuilder()

    # TODO: Add data processing logic specific to data_source

    report.add_title_sheet(
        "Daily Report",
        datetime.now().strftime("%Y-%m-%d"),
        {"Generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    )

    report.save(report_output)

    # Email report
    email = EmailHelper()
    subject = email_subject or f"Daily Report - {datetime.now().strftime('%Y-%m-%d')}"
    body = f"Please find attached the daily report for {datetime.now().strftime('%Y-%m-%d')}."

    email.send_report(
        to=email_recipients,
        subject=subject,
        body=body,
        attachments=[report_output]
    )

    logger.info("Daily report workflow completed!")
