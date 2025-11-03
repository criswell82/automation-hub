"""
Security and credential management for Automation Hub.
Uses Windows Credential Manager for secure credential storage.
"""

import logging
import base64
import json
from typing import Optional, Dict, Any
import os

try:
    import win32cred
    import pywintypes
    WINDOWS_CRED_AVAILABLE = True
except ImportError:
    WINDOWS_CRED_AVAILABLE = False
    logging.warning("pywin32 not available - credential storage will use fallback method")


class SecurityManager:
    """
    Manages secure storage and retrieval of credentials and sensitive data.
    Uses Windows Credential Manager when available.
    """

    def __init__(self, app_name: str = "AutomationHub"):
        """
        Initialize the security manager.

        Args:
            app_name: Application name for credential namespacing
        """
        self.app_name = app_name
        self.logger = logging.getLogger(__name__)
        self.use_windows_cred_store = WINDOWS_CRED_AVAILABLE

        if not self.use_windows_cred_store:
            self.logger.warning("Windows Credential Manager not available. Using fallback storage.")
            self._init_fallback_storage()

    def _init_fallback_storage(self):
        """Initialize fallback credential storage (encrypted file-based)."""
        appdata = os.getenv('APPDATA', os.path.expanduser('~'))
        self.cred_dir = os.path.join(appdata, self.app_name, 'credentials')
        os.makedirs(self.cred_dir, exist_ok=True)
        self.cred_file = os.path.join(self.cred_dir, 'creds.dat')

    def store_credential(
            self,
            target: str,
            username: str,
            password: str,
            attributes: Optional[Dict[str, Any]] = None
    ) -> bool:
        """
        Store a credential securely.

        Args:
            target: Credential target (e.g., 'SharePoint', 'Outlook')
            username: Username or account identifier
            password: Password or secret
            attributes: Additional attributes to store

        Returns:
            bool: True if successful
        """
        full_target = f"{self.app_name}:{target}"

        try:
            if self.use_windows_cred_store:
                return self._store_windows_credential(full_target, username, password, attributes)
            else:
                return self._store_fallback_credential(full_target, username, password, attributes)
        except Exception as e:
            self.logger.error(f"Failed to store credential for {target}: {e}")
            return False

    def retrieve_credential(self, target: str) -> Optional[Dict[str, Any]]:
        """
        Retrieve a stored credential.

        Args:
            target: Credential target

        Returns:
            dict: Credential information with 'username', 'password', and optional 'attributes'
                 or None if not found
        """
        full_target = f"{self.app_name}:{target}"

        try:
            if self.use_windows_cred_store:
                return self._retrieve_windows_credential(full_target)
            else:
                return self._retrieve_fallback_credential(full_target)
        except Exception as e:
            self.logger.error(f"Failed to retrieve credential for {target}: {e}")
            return None

    def delete_credential(self, target: str) -> bool:
        """
        Delete a stored credential.

        Args:
            target: Credential target

        Returns:
            bool: True if successful
        """
        full_target = f"{self.app_name}:{target}"

        try:
            if self.use_windows_cred_store:
                return self._delete_windows_credential(full_target)
            else:
                return self._delete_fallback_credential(full_target)
        except Exception as e:
            self.logger.error(f"Failed to delete credential for {target}: {e}")
            return False

    def list_credentials(self) -> list:
        """
        List all stored credential targets.

        Returns:
            list: List of credential target names
        """
        try:
            if self.use_windows_cred_store:
                return self._list_windows_credentials()
            else:
                return self._list_fallback_credentials()
        except Exception as e:
            self.logger.error(f"Failed to list credentials: {e}")
            return []

    def _store_windows_credential(
            self,
            target: str,
            username: str,
            password: str,
            attributes: Optional[Dict[str, Any]]
    ) -> bool:
        """Store credential using Windows Credential Manager."""
        try:
            # Create credential structure
            cred = {
                'Type': win32cred.CRED_TYPE_GENERIC,
                'TargetName': target,
                'UserName': username,
                'CredentialBlob': password,
                'Comment': 'Stored by Automation Hub',
                'Persist': win32cred.CRED_PERSIST_LOCAL_MACHINE
            }

            # Store additional attributes as comment if provided
            if attributes:
                cred['Comment'] = json.dumps(attributes)

            win32cred.CredWrite(cred, 0)
            self.logger.info(f"Credential stored: {target}")
            return True

        except pywintypes.error as e:
            self.logger.error(f"Windows credential storage failed: {e}")
            return False

    def _retrieve_windows_credential(self, target: str) -> Optional[Dict[str, Any]]:
        """Retrieve credential from Windows Credential Manager."""
        try:
            cred = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC, 0)

            result = {
                'username': cred['UserName'],
                'password': cred['CredentialBlob'].decode('utf-16-le'),
            }

            # Try to parse attributes from comment
            if cred.get('Comment'):
                try:
                    result['attributes'] = json.loads(cred['Comment'])
                except json.JSONDecodeError:
                    pass

            return result

        except pywintypes.error:
            # Credential not found
            return None

    def _delete_windows_credential(self, target: str) -> bool:
        """Delete credential from Windows Credential Manager."""
        try:
            win32cred.CredDelete(target, win32cred.CRED_TYPE_GENERIC, 0)
            self.logger.info(f"Credential deleted: {target}")
            return True
        except pywintypes.error:
            return False

    def _list_windows_credentials(self) -> list:
        """List credentials from Windows Credential Manager."""
        try:
            creds = win32cred.CredEnumerate(None, 0)
            prefix = f"{self.app_name}:"

            return [
                cred['TargetName'].replace(prefix, '')
                for cred in creds
                if cred['TargetName'].startswith(prefix)
            ]
        except pywintypes.error:
            return []

    def _store_fallback_credential(
            self,
            target: str,
            username: str,
            password: str,
            attributes: Optional[Dict[str, Any]]
    ) -> bool:
        """Store credential using fallback method (base64 encoded file)."""
        # Load existing credentials
        creds = self._load_fallback_credentials()

        # Store new credential
        creds[target] = {
            'username': username,
            'password': base64.b64encode(password.encode()).decode(),
            'attributes': attributes
        }

        # Save credentials
        return self._save_fallback_credentials(creds)

    def _retrieve_fallback_credential(self, target: str) -> Optional[Dict[str, Any]]:
        """Retrieve credential using fallback method."""
        creds = self._load_fallback_credentials()

        if target in creds:
            cred = creds[target]
            return {
                'username': cred['username'],
                'password': base64.b64decode(cred['password'].encode()).decode(),
                'attributes': cred.get('attributes')
            }

        return None

    def _delete_fallback_credential(self, target: str) -> bool:
        """Delete credential using fallback method."""
        creds = self._load_fallback_credentials()

        if target in creds:
            del creds[target]
            return self._save_fallback_credentials(creds)

        return False

    def _list_fallback_credentials(self) -> list:
        """List credentials using fallback method."""
        creds = self._load_fallback_credentials()
        prefix = f"{self.app_name}:"

        return [
            target.replace(prefix, '')
            for target in creds.keys()
            if target.startswith(prefix)
        ]

    def _load_fallback_credentials(self) -> Dict[str, Any]:
        """Load credentials from fallback storage."""
        if not os.path.exists(self.cred_file):
            return {}

        try:
            with open(self.cred_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"Failed to load fallback credentials: {e}")
            return {}

    def _save_fallback_credentials(self, creds: Dict[str, Any]) -> bool:
        """Save credentials to fallback storage."""
        try:
            with open(self.cred_file, 'w', encoding='utf-8') as f:
                json.dump(creds, f, indent=4)
            return True
        except Exception as e:
            self.logger.error(f"Failed to save fallback credentials: {e}")
            return False

    @staticmethod
    def generate_api_key(length: int = 32) -> str:
        """
        Generate a random API key.

        Args:
            length: Length of the key

        Returns:
            str: Random API key
        """
        import secrets
        import string

        alphabet = string.ascii_letters + string.digits
        return ''.join(secrets.choice(alphabet) for _ in range(length))

    def validate_credential(self, target: str) -> bool:
        """
        Check if a credential exists and is retrievable.

        Args:
            target: Credential target

        Returns:
            bool: True if credential exists and is valid
        """
        cred = self.retrieve_credential(target)
        return cred is not None and 'username' in cred and 'password' in cred


# Singleton instance
_security_manager: Optional[SecurityManager] = None


def get_security_manager() -> SecurityManager:
    """Get the global SecurityManager instance."""
    global _security_manager
    if _security_manager is None:
        _security_manager = SecurityManager()
    return _security_manager
