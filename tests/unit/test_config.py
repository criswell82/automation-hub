"""
Unit tests for ConfigManager.
"""

import json
import pytest
from pathlib import Path

from core.config import ConfigManager


class TestConfigManager:
    """Tests for ConfigManager class."""

    def test_init_creates_config_dir(self, temp_config_dir: Path):
        """Test that ConfigManager creates config directory on init."""
        config_dir = temp_config_dir / "new_config"
        manager = ConfigManager(config_dir=str(config_dir))

        assert config_dir.exists()
        assert manager.config_dir == config_dir

    def test_default_config_loaded(self, temp_config_dir: Path):
        """Test that default configuration is loaded."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        assert manager.get('version') == '1.0.0'
        assert manager.get('app.theme') is not None
        assert manager.get('scheduler.max_concurrent_jobs') is not None

    def test_get_with_dot_notation(self, temp_config_dir: Path):
        """Test getting nested config values with dot notation."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        theme = manager.get('app.theme')
        assert theme is not None

        max_jobs = manager.get('scheduler.max_concurrent_jobs')
        assert isinstance(max_jobs, int)

    def test_get_returns_default_for_missing_key(self, temp_config_dir: Path):
        """Test that get() returns default for missing keys."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        result = manager.get('nonexistent.key', 'default_value')
        assert result == 'default_value'

        result = manager.get('also.missing')
        assert result is None

    def test_set_creates_nested_keys(self, temp_config_dir: Path):
        """Test that set() creates nested dictionary structure."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        manager.set('new.nested.key', 'test_value', save=False)

        assert manager.get('new.nested.key') == 'test_value'

    def test_set_and_save(self, temp_config_dir: Path):
        """Test that set() saves to file when save=True."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        manager.set('app.theme', 'custom_theme')

        # Reload and verify persistence
        manager2 = ConfigManager(config_dir=str(temp_config_dir))
        assert manager2.get('app.theme') == 'custom_theme'

    def test_module_config(self, temp_config_dir: Path):
        """Test module-specific configuration."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        # Set module config
        module_config = {'setting1': 'value1', 'setting2': 42}
        manager.set_module_config('excel', module_config)

        # Retrieve module config
        retrieved = manager.get_module_config('excel')
        assert retrieved == module_config

    def test_get_module_config_default(self, temp_config_dir: Path):
        """Test get_module_config returns default for missing module."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        default = {'default': 'config'}
        result = manager.get_module_config('nonexistent', default=default)
        assert result == default

    def test_get_all(self, temp_config_dir: Path):
        """Test get_all returns copy of config."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        all_config = manager.get_all()

        assert isinstance(all_config, dict)
        assert 'version' in all_config
        assert 'app' in all_config

    def test_reset_to_defaults(self, temp_config_dir: Path):
        """Test reset_to_defaults restores default configuration."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        # Modify config
        manager.set('app.theme', 'modified_theme')

        # Reset
        manager.reset_to_defaults()

        # Verify default is restored
        assert manager.get('app.theme') == 'light'

    def test_export_json(self, temp_config_dir: Path, temp_dir: Path):
        """Test exporting configuration to JSON."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        export_path = temp_dir / "exported_config.json"
        manager.export_config(str(export_path), format='json')

        assert export_path.exists()

        with open(export_path) as f:
            exported = json.load(f)

        assert exported['version'] == manager.get('version')

    def test_export_yaml(self, temp_config_dir: Path, temp_dir: Path):
        """Test exporting configuration to YAML."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        export_path = temp_dir / "exported_config.yaml"
        manager.export_config(str(export_path), format='yaml')

        assert export_path.exists()

    def test_import_json(self, temp_config_dir: Path, temp_dir: Path, sample_config: dict):
        """Test importing configuration from JSON."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        # Create import file
        import_path = temp_dir / "import_config.json"
        with open(import_path, 'w') as f:
            json.dump(sample_config, f)

        # Import
        manager.import_config(str(import_path), format='json', merge=True)

        assert manager.get('app.theme') == 'dark'
        assert manager.get('scheduler.max_concurrent_jobs') == 3

    def test_import_nonexistent_file_raises(self, temp_config_dir: Path):
        """Test that importing nonexistent file raises error."""
        manager = ConfigManager(config_dir=str(temp_config_dir))

        with pytest.raises(FileNotFoundError):
            manager.import_config('/nonexistent/path.json')

    def test_deep_merge(self):
        """Test deep merge of dictionaries."""
        base = {
            'level1': {
                'level2': {
                    'key1': 'base_value1',
                    'key2': 'base_value2'
                }
            }
        }

        overlay = {
            'level1': {
                'level2': {
                    'key2': 'overlay_value2',
                    'key3': 'overlay_value3'
                }
            }
        }

        result = ConfigManager._deep_merge(base, overlay)

        assert result['level1']['level2']['key1'] == 'base_value1'
        assert result['level1']['level2']['key2'] == 'overlay_value2'
        assert result['level1']['level2']['key3'] == 'overlay_value3'
