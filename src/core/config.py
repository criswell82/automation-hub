"""
Configuration management system for Automation Hub.
Handles application and module-specific configuration using JSON and YAML formats.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional
import yaml


class ConfigManager:
    """
    Manages application configuration with support for multiple config files,
    defaults, and environment-specific settings.
    """

    def __init__(self, config_dir: Optional[str] = None):
        """
        Initialize the configuration manager.

        Args:
            config_dir: Directory to store configuration files.
                       Defaults to user's AppData/AutomationHub/config
        """
        if config_dir is None:
            # Default to user's AppData directory
            appdata = os.getenv('APPDATA', os.path.expanduser('~'))
            config_dir = os.path.join(appdata, 'AutomationHub', 'config')

        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)

        self.config_file = self.config_dir / 'config.json'
        self.module_config_file = self.config_dir / 'modules.json'

        self._config: Dict[str, Any] = {}
        self._module_configs: Dict[str, Dict[str, Any]] = {}

        self._load_default_config()
        self.load()

    def _load_default_config(self):
        """Load default configuration values."""
        self._config = {
            'version': '1.0.0',
            'app': {
                'theme': 'light',
                'language': 'en',
                'check_updates': True,
                'auto_update': False,
                'log_level': 'INFO',
                'max_recent_scripts': 10
            },
            'paths': {
                'logs': str(self.config_dir.parent / 'logs'),
                'temp': str(self.config_dir.parent / 'temp'),
                'scripts': str(self.config_dir.parent / 'scripts'),
                'output': str(self.config_dir.parent / 'output')
            },
            'scheduler': {
                'max_concurrent_jobs': 5,
                'job_timeout_minutes': 60,
                'retry_failed_jobs': True,
                'max_retries': 3
            },
            'security': {
                'use_windows_credential_store': True,
                'encrypt_config': False,
                'require_authentication': False
            }
        }

        # Ensure directories exist
        for path_key, path_value in self._config['paths'].items():
            Path(path_value).mkdir(parents=True, exist_ok=True)

    def load(self):
        """Load configuration from file."""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults (loaded config takes precedence)
                    self._deep_merge(self._config, loaded_config)

            if self.module_config_file.exists():
                with open(self.module_config_file, 'r', encoding='utf-8') as f:
                    self._module_configs = json.load(f)

        except Exception as e:
            print(f"Warning: Could not load configuration: {e}")
            # Continue with default config

    def save(self):
        """Save current configuration to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=4)

            with open(self.module_config_file, 'w', encoding='utf-8') as f:
                json.dump(self._module_configs, f, indent=4)

        except Exception as e:
            raise RuntimeError(f"Failed to save configuration: {e}")

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get a configuration value using dot notation.

        Args:
            key: Configuration key (e.g., 'app.theme' or 'scheduler.max_concurrent_jobs')
            default: Default value if key not found

        Returns:
            Configuration value or default
        """
        keys = key.split('.')
        value = self._config

        try:
            for k in keys:
                value = value[k]
            return value
        except (KeyError, TypeError):
            return default

    def set(self, key: str, value: Any, save: bool = True):
        """
        Set a configuration value using dot notation.

        Args:
            key: Configuration key (e.g., 'app.theme')
            value: Value to set
            save: Whether to immediately save to file
        """
        keys = key.split('.')
        config = self._config

        # Navigate to the correct nested dict
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]

        # Set the value
        config[keys[-1]] = value

        if save:
            self.save()

    def get_module_config(self, module_name: str, default: Optional[Dict] = None) -> Dict[str, Any]:
        """
        Get configuration for a specific module.

        Args:
            module_name: Name of the module
            default: Default configuration if module not found

        Returns:
            Module configuration dictionary
        """
        return self._module_configs.get(module_name, default or {})

    def set_module_config(self, module_name: str, config: Dict[str, Any], save: bool = True):
        """
        Set configuration for a specific module.

        Args:
            module_name: Name of the module
            config: Configuration dictionary
            save: Whether to immediately save to file
        """
        self._module_configs[module_name] = config

        if save:
            self.save()

    def get_all(self) -> Dict[str, Any]:
        """Get the entire configuration dictionary."""
        return self._config.copy()

    def get_all_module_configs(self) -> Dict[str, Dict[str, Any]]:
        """Get all module configurations."""
        return self._module_configs.copy()

    def reset_to_defaults(self):
        """Reset configuration to default values."""
        self._config = {}
        self._module_configs = {}
        self._load_default_config()
        self.save()

    def export_config(self, filepath: str, format: str = 'json'):
        """
        Export configuration to a file.

        Args:
            filepath: Path to export file
            format: Export format ('json' or 'yaml')
        """
        filepath = Path(filepath)

        try:
            if format == 'yaml':
                with open(filepath, 'w', encoding='utf-8') as f:
                    yaml.dump(self._config, f, default_flow_style=False)
            else:  # json
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(self._config, f, indent=4)
        except Exception as e:
            raise RuntimeError(f"Failed to export configuration: {e}")

    def import_config(self, filepath: str, format: str = 'json', merge: bool = True):
        """
        Import configuration from a file.

        Args:
            filepath: Path to import file
            format: Import format ('json' or 'yaml')
            merge: If True, merge with existing config; if False, replace
        """
        filepath = Path(filepath)

        if not filepath.exists():
            raise FileNotFoundError(f"Configuration file not found: {filepath}")

        try:
            if format == 'yaml':
                with open(filepath, 'r', encoding='utf-8') as f:
                    imported_config = yaml.safe_load(f)
            else:  # json
                with open(filepath, 'r', encoding='utf-8') as f:
                    imported_config = json.load(f)

            if merge:
                self._deep_merge(self._config, imported_config)
            else:
                self._config = imported_config

            self.save()
        except Exception as e:
            raise RuntimeError(f"Failed to import configuration: {e}")

    @staticmethod
    def _deep_merge(base: Dict, overlay: Dict) -> Dict:
        """
        Deep merge two dictionaries. Overlay values take precedence.

        Args:
            base: Base dictionary
            overlay: Overlay dictionary

        Returns:
            Merged dictionary
        """
        for key, value in overlay.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                ConfigManager._deep_merge(base[key], value)
            else:
                base[key] = value
        return base


# Singleton instance
_config_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """Get the global ConfigManager instance."""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager
