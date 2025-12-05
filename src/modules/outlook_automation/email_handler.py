"""Email Handler for Outlook automation using win32com."""

import logging
from typing import Any, Dict, List, Optional

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


class EmailHandler:
    """Handle Outlook email operations."""

    def __init__(self) -> None:
        """Initialize the email handler."""
        if not WIN32COM_AVAILABLE:
            raise RuntimeError("pywin32 is required for Outlook automation")

        self.logger = logging.getLogger(__name__)
        self.outlook: Optional[Any] = None

    def connect(self) -> None:
        """Connect to Outlook application."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.logger.info("Connected to Outlook")
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            raise

    def send_email(
        self,
        to: str,
        subject: str,
        body: str,
        attachments: Optional[List[str]] = None,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        html_body: bool = False
    ) -> None:
        """
        Send an email.

        Args:
            to: Recipient email address(es), semicolon-separated for multiple.
            subject: Email subject.
            body: Email body content.
            attachments: List of file paths to attach.
            cc: CC recipients, semicolon-separated.
            bcc: BCC recipients, semicolon-separated.
            html_body: If True, treat body as HTML content.
        """
        if not self.outlook:
            self.connect()

        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject

        if html_body:
            mail.HTMLBody = body
        else:
            mail.Body = body

        if cc:
            mail.CC = cc

        if bcc:
            mail.BCC = bcc

        if attachments:
            for attachment in attachments:
                mail.Attachments.Add(attachment)

        mail.Send()
        self.logger.info(f"Sent email to {to}")

    def read_emails(
        self,
        folder: str = "Inbox",
        unread_only: bool = False,
        count: int = 10
    ) -> List[Dict[str, str]]:
        """
        Read emails from a folder.

        Args:
            folder: Folder name to read from.
            unread_only: If True, only return unread emails.
            count: Maximum number of emails to return.

        Returns:
            List of email dictionaries with subject, sender, received, body.
        """
        if not self.outlook:
            self.connect()

        namespace = self.outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) if folder == "Inbox" else namespace.Folders[folder]

        messages = inbox.Items
        if unread_only:
            messages = messages.Restrict("[Unread]=True")

        emails: List[Dict[str, str]] = []
        for i, message in enumerate(messages):
            if i >= count:
                break
            emails.append({
                'subject': message.Subject,
                'sender': message.SenderName,
                'received': str(message.ReceivedTime),
                'body': message.Body
            })

        return emails
