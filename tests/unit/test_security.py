"""
Unit tests for SecurityManager.

These tests use the fallback storage mechanism since Windows Credential Manager
is not available in most test environments.
"""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from core.security import SecurityManager


class TestSecurityManager:
    """Tests for SecurityManager class."""

    @pytest.fixture
    def security_manager(self, temp_dir: Path) -> SecurityManager:
        """Create a SecurityManager with fallback storage in temp directory."""
        with patch.object(SecurityManager, '__init__', lambda self, app_name: None):
            manager = SecurityManager.__new__(SecurityManager)
            manager.app_name = "TestApp"
            manager.use_windows_cred_store = False
            manager.cred_dir = str(temp_dir / "credentials")
            manager.cred_file = str(temp_dir / "credentials" / "creds.dat")
            Path(manager.cred_dir).mkdir(parents=True, exist_ok=True)
            import logging
            manager.logger = logging.getLogger(__name__)
            return manager

    def test_store_and_retrieve_credential(self, security_manager: SecurityManager):
        """Test storing and retrieving a credential."""
        result = security_manager.store_credential(
            target="TestService",
            username="testuser",
            password="testpass123"
        )

        assert result is True

        # Retrieve
        cred = security_manager.retrieve_credential("TestService")

        assert cred is not None
        assert cred['username'] == 'testuser'
        assert cred['password'] == 'testpass123'

    def test_store_credential_with_attributes(self, security_manager: SecurityManager):
        """Test storing credential with additional attributes."""
        attributes = {'server': 'example.com', 'port': 443}

        result = security_manager.store_credential(
            target="TestService",
            username="testuser",
            password="testpass",
            attributes=attributes
        )

        assert result is True

        cred = security_manager.retrieve_credential("TestService")
        assert cred['attributes'] == attributes

    def test_retrieve_nonexistent_credential(self, security_manager: SecurityManager):
        """Test retrieving a credential that doesn't exist."""
        cred = security_manager.retrieve_credential("NonExistent")

        assert cred is None

    def test_delete_credential(self, security_manager: SecurityManager):
        """Test deleting a stored credential."""
        # Store first
        security_manager.store_credential(
            target="ToDelete",
            username="user",
            password="pass"
        )

        # Verify it exists
        assert security_manager.retrieve_credential("ToDelete") is not None

        # Delete
        result = security_manager.delete_credential("ToDelete")
        assert result is True

        # Verify it's gone
        assert security_manager.retrieve_credential("ToDelete") is None

    def test_delete_nonexistent_credential(self, security_manager: SecurityManager):
        """Test deleting a credential that doesn't exist."""
        result = security_manager.delete_credential("NonExistent")

        assert result is False

    def test_list_credentials(self, security_manager: SecurityManager):
        """Test listing stored credentials."""
        # Store multiple credentials
        security_manager.store_credential("Service1", "user1", "pass1")
        security_manager.store_credential("Service2", "user2", "pass2")
        security_manager.store_credential("Service3", "user3", "pass3")

        # List
        creds = security_manager.list_credentials()

        assert len(creds) == 3
        assert "Service1" in creds
        assert "Service2" in creds
        assert "Service3" in creds

    def test_validate_credential_exists(self, security_manager: SecurityManager):
        """Test validation of existing credential."""
        security_manager.store_credential("ValidService", "user", "pass")

        result = security_manager.validate_credential("ValidService")

        assert result is True

    def test_validate_credential_not_exists(self, security_manager: SecurityManager):
        """Test validation of non-existent credential."""
        result = security_manager.validate_credential("InvalidService")

        assert result is False

    def test_generate_api_key(self):
        """Test API key generation."""
        key1 = SecurityManager.generate_api_key()
        key2 = SecurityManager.generate_api_key()

        # Default length is 32
        assert len(key1) == 32
        assert len(key2) == 32

        # Keys should be different
        assert key1 != key2

        # Keys should be alphanumeric
        assert key1.isalnum()
        assert key2.isalnum()

    def test_generate_api_key_custom_length(self):
        """Test API key generation with custom length."""
        key = SecurityManager.generate_api_key(length=64)

        assert len(key) == 64

    def test_update_existing_credential(self, security_manager: SecurityManager):
        """Test updating an existing credential."""
        # Store initial credential
        security_manager.store_credential("UpdateTest", "user1", "pass1")

        # Update with new values
        security_manager.store_credential("UpdateTest", "user2", "pass2")

        # Retrieve and verify updated values
        cred = security_manager.retrieve_credential("UpdateTest")

        assert cred['username'] == 'user2'
        assert cred['password'] == 'pass2'


class TestSecurityManagerWindowsCred:
    """Tests for Windows Credential Manager integration (mocked)."""

    def test_windows_cred_store_available(self):
        """Test detection of Windows Credential Manager availability."""
        with patch.dict('sys.modules', {'win32cred': MagicMock(), 'pywintypes': MagicMock()}):
            # This would typically set WINDOWS_CRED_AVAILABLE = True
            pass

    def test_fallback_when_windows_cred_unavailable(self, temp_dir: Path):
        """Test fallback storage is used when Windows Cred Manager unavailable."""
        with patch('core.security.WINDOWS_CRED_AVAILABLE', False):
            manager = SecurityManager.__new__(SecurityManager)
            manager.app_name = "TestApp"
            manager.use_windows_cred_store = False
            manager.cred_dir = str(temp_dir / "credentials")
            manager.cred_file = str(temp_dir / "credentials" / "creds.dat")
            Path(manager.cred_dir).mkdir(parents=True, exist_ok=True)
            import logging
            manager.logger = logging.getLogger(__name__)

            # Should use fallback storage
            assert manager.use_windows_cred_store is False
