"""Email Handler for Outlook automation using win32com."""

import logging

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


class EmailHandler:
    """Handle Outlook email operations."""

    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise RuntimeError("pywin32 is required for Outlook automation")

        self.logger = logging.getLogger(__name__)
        self.outlook = None

    def connect(self):
        """Connect to Outlook."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.logger.info("Connected to Outlook")
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            raise

    def send_email(self, to: str, subject: str, body: str, attachments: list = None):
        """Send an email."""
        if not self.outlook:
            self.connect()

        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.Body = body

        if attachments:
            for attachment in attachments:
                mail.Attachments.Add(attachment)

        mail.Send()
        self.logger.info(f"Sent email to {to}")

    def read_emails(self, folder: str = "Inbox", unread_only: bool = False, count: int = 10):
        """Read emails from a folder."""
        if not self.outlook:
            self.connect()

        namespace = self.outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) if folder == "Inbox" else namespace.Folders[folder]

        messages = inbox.Items
        if unread_only:
            messages = messages.Restrict("[Unread]=True")

        emails = []
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
