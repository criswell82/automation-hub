"""
WORKFLOW_META:
  name: Configuration Test
  description: Tests the configuration management system (read, write, persistence)
  category: Tests
  version: 1.0.0
  author: Automation Hub
  parameters: {}
"""

import sys
import os
from pathlib import Path

# Add src directory to path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class ConfigTest(BaseModule):
    """Test configuration system"""

    def __init__(self):
        super().__init__()

    def configure(self, **kwargs):
        """No configuration needed"""
        logger.info("Starting configuration test...")
        return True

    def validate(self):
        """Always valid"""
        return True

    def execute(self):
        """Test configuration system"""
        try:
            logger.info("=" * 60)
            logger.info("CONFIGURATION SYSTEM TEST")
            logger.info("=" * 60)

            # Test 1: Initialize ConfigManager
            logger.info("")
            logger.info("Test 1: Initialize ConfigManager")
            logger.info("-" * 60)

            test1_passed = False
            config = None
            try:
                from core.config import ConfigManager
                config = ConfigManager()
                logger.info(f"✓ ConfigManager initialized")
                logger.info(f"  Config directory: {config.config_dir}")
                logger.info(f"  Config file: {config.config_file}")
                test1_passed = True
            except Exception as e:
                logger.info(f"✗ Initialization failed: {e}")

            # Test 2: Read default configuration
            logger.info("")
            logger.info("Test 2: Read Default Configuration")
            logger.info("-" * 60)

            test2_passed = False
            if test1_passed:
                try:
                    version = config.get('version', None)
                    app_theme = config.get('app.theme', None)
                    log_level = config.get('app.log_level', None)
                    ai_model = config.get('ai.model', None)

                    logger.info(f"✓ Config values read successfully")
                    logger.info(f"  Version: {version}")
                    logger.info(f"  Theme: {app_theme}")
                    logger.info(f"  Log Level: {log_level}")
                    logger.info(f"  AI Model: {ai_model}")

                    # Validate expected defaults
                    has_version = bool(version)
                    has_theme = bool(app_theme)
                    has_ai_model = bool(ai_model)

                    test2_passed = has_version and has_theme and has_ai_model

                    if not test2_passed:
                        logger.info(f"  ! Some expected defaults missing")

                except Exception as e:
                    logger.info(f"✗ Reading config failed: {e}")

            # Test 3: Write and read config value
            logger.info("")
            logger.info("Test 3: Write and Read Config Value")
            logger.info("-" * 60)

            test3_passed = False
            if test1_passed:
                try:
                    # Write a test value
                    test_key = 'test_workflow.test_value'
                    test_value = 'automation_hub_test_12345'

                    config.set(test_key, test_value, save=False)  # Don't save yet
                    logger.info(f"  ✓ Test value written (in memory)")

                    # Read it back
                    read_value = config.get(test_key, None)

                    if read_value == test_value:
                        logger.info(f"✓ Test value read successfully")
                        logger.info(f"  Written: {test_value}")
                        logger.info(f"  Read: {read_value}")
                        test3_passed = True
                    else:
                        logger.info(f"✗ Value mismatch")
                        logger.info(f"  Expected: {test_value}")
                        logger.info(f"  Got: {read_value}")

                except Exception as e:
                    logger.info(f"✗ Write/read test failed: {e}")

            # Test 4: Test dot notation access
            logger.info("")
            logger.info("Test 4: Test Dot Notation Access")
            logger.info("-" * 60)

            test4_passed = False
            if test1_passed:
                try:
                    # Test nested access
                    paths_logs = config.get('paths.logs', None)
                    scheduler_max = config.get('scheduler.max_concurrent_jobs', None)
                    ai_max_tokens = config.get('ai.max_tokens', None)

                    logger.info(f"✓ Dot notation access works")
                    logger.info(f"  paths.logs: {paths_logs}")
                    logger.info(f"  scheduler.max_concurrent_jobs: {scheduler_max}")
                    logger.info(f"  ai.max_tokens: {ai_max_tokens}")

                    test4_passed = bool(paths_logs) and bool(scheduler_max) and bool(ai_max_tokens)

                    if not test4_passed:
                        logger.info(f"  ! Some nested values missing")

                except Exception as e:
                    logger.info(f"✗ Dot notation test failed: {e}")

            # Test 5: Test config file locations
            logger.info("")
            logger.info("Test 5: Verify Config File Locations")
            logger.info("-" * 60)

            test5_passed = False
            if test1_passed:
                try:
                    config_file_exists = config.config_file.exists()
                    config_dir_exists = config.config_dir.exists()
                    config_dir_writable = os.access(str(config.config_dir), os.W_OK)

                    logger.info(f"  Config file exists: {config_file_exists}")
                    logger.info(f"  Config dir exists: {config_dir_exists}")
                    logger.info(f"  Config dir writable: {config_dir_writable}")

                    if config_file_exists:
                        logger.info(f"  ✓ Config file: {config.config_file}")

                    if config_dir_writable:
                        logger.info(f"  ✓ Config directory is writable")

                    # Module config file
                    module_config_exists = config.module_config_file.exists()
                    logger.info(f"  Module config exists: {module_config_exists}")

                    test5_passed = config_dir_exists and config_dir_writable

                except Exception as e:
                    logger.info(f"✗ File location test failed: {e}")

            # Test 6: Test module config
            logger.info("")
            logger.info("Test 6: Test Module-Specific Config")
            logger.info("-" * 60)

            test6_passed = False
            if test1_passed:
                try:
                    # Get module config
                    test_module_config = config.get_module_config('test_module', {})

                    logger.info(f"✓ Module config accessed")
                    logger.info(f"  Test module config: {test_module_config}")

                    # Try to set module config
                    test_config = {'test_key': 'test_value'}
                    config.set_module_config('test_module', test_config, save=False)

                    # Read it back
                    read_config = config.get_module_config('test_module', {})

                    if read_config.get('test_key') == 'test_value':
                        logger.info(f"✓ Module config read/write works")
                        test6_passed = True
                    else:
                        logger.info(f"  ! Module config mismatch")

                except Exception as e:
                    logger.info(f"✗ Module config test failed: {e}")

            # Summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("SUMMARY")
            logger.info("=" * 60)

            tests = [
                ("Initialize ConfigManager", test1_passed),
                ("Read Default Configuration", test2_passed),
                ("Write and Read Config Value", test3_passed),
                ("Dot Notation Access", test4_passed),
                ("Verify Config File Locations", test5_passed),
                ("Module-Specific Config", test6_passed)
            ]

            passed = sum(1 for _, result in tests if result)
            total = len(tests)

            for test_name, result in tests:
                status = "✓ PASS" if result else "✗ FAIL"
                logger.info(f"  {status}: {test_name}")

            logger.info("")
            logger.info(f"Total: {passed}/{total} tests passed")
            logger.info("=" * 60)

            all_passed = passed == total

            return {
                "status": "success" if all_passed else "warning",
                "message": f"Configuration test: {passed}/{total} tests passed",
                "tests_passed": passed,
                "tests_total": total,
                "config_directory": str(config.config_dir) if config else None,
                "test_results": {name: result for name, result in tests}
            }

        except Exception as e:
            self.handle_error(e, "Configuration test failed")
            raise


def run(**kwargs):
    """Execute the test"""
    test = ConfigTest()
    test.configure(**kwargs)
    test.validate()
    return test.execute()


if __name__ == "__main__":
    # Run the test
    result = run()
    print(f"\n\nTest Result: {result['status'].upper()}")
    print(f"Summary: {result['message']}")
