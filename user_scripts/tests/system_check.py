"""
WORKFLOW_META:
  name: System Check Test
  description: Validates Automation Hub system is correctly configured and all dependencies are installed
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


class SystemCheckTest(BaseModule):
    """System validation test workflow"""

    def __init__(self):
        super().__init__()
        self.results = {}

    def configure(self, **kwargs):
        """No configuration needed"""
        logger.info("Starting system check...")
        return True

    def validate(self):
        """Always valid"""
        return True

    def execute(self):
        """Run all system checks"""
        try:
            logger.info("=" * 60)
            logger.info("AUTOMATION HUB SYSTEM CHECK")
            logger.info("=" * 60)

            # Run all checks
            self.results = {
                "dependencies": self._check_dependencies(),
                "directories": self._check_directories(),
                "config_system": self._check_config(),
                "security_system": self._check_security(),
                "modules": self._check_modules()
            }

            # Calculate summary
            total_checks = sum(r["checks_run"] for r in self.results.values())
            passed_checks = sum(r["checks_passed"] for r in self.results.values())
            failed_checks = total_checks - passed_checks

            # Print summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("SUMMARY")
            logger.info("=" * 60)
            logger.info(f"Total Checks: {total_checks}")
            logger.info(f"✓ Passed: {passed_checks}")
            logger.info(f"✗ Failed: {failed_checks}")
            logger.info("=" * 60)

            all_passed = failed_checks == 0

            return {
                "status": "success" if all_passed else "warning",
                "message": f"System check completed: {passed_checks}/{total_checks} checks passed",
                "results": self.results,
                "total_checks": total_checks,
                "passed": passed_checks,
                "failed": failed_checks
            }

        except Exception as e:
            self.handle_error(e, "System check failed")
            raise

    def _check_dependencies(self):
        """Check if all required dependencies are installed"""
        logger.info("")
        logger.info("Checking Dependencies...")
        logger.info("-" * 60)

        dependencies = {
            'PyQt5': 'PyQt5',
            'anthropic': 'anthropic',
            'openpyxl': 'openpyxl',
            'pywin32': 'win32api',
            'python-docx': 'docx',
            'APScheduler': 'apscheduler',
            'PyYAML': 'yaml',
            'Pillow': 'PIL',
            'pyautogui': 'pyautogui',
            'requests': 'requests'
        }

        passed = 0
        failed = 0

        for package_name, import_name in dependencies.items():
            try:
                __import__(import_name)
                logger.info(f"  ✓ {package_name}: INSTALLED")
                passed += 1
            except ImportError:
                logger.info(f"  ✗ {package_name}: MISSING")
                failed += 1

        return {
            "status": "PASS" if failed == 0 else "FAIL",
            "checks_run": len(dependencies),
            "checks_passed": passed,
            "details": f"{passed}/{len(dependencies)} dependencies installed"
        }

    def _check_directories(self):
        """Check if required directories exist"""
        logger.info("")
        logger.info("Checking Directory Structure...")
        logger.info("-" * 60)

        appdata = os.getenv('APPDATA', os.path.expanduser('~'))
        base_dir = os.path.join(appdata, 'AutomationHub')

        directories = {
            'Config': os.path.join(base_dir, 'config'),
            'Logs': os.path.join(base_dir, 'logs'),
            'Temp': os.path.join(base_dir, 'temp'),
            'User Scripts - Custom': str(project_root / 'user_scripts' / 'custom'),
            'User Scripts - Templates': str(project_root / 'user_scripts' / 'templates')
        }

        passed = 0
        failed = 0

        for name, path in directories.items():
            if os.path.exists(path):
                logger.info(f"  ✓ {name}: {path}")
                passed += 1
            else:
                logger.info(f"  ✗ {name}: MISSING - {path}")
                failed += 1

        return {
            "status": "PASS" if failed == 0 else "FAIL",
            "checks_run": len(directories),
            "checks_passed": passed,
            "details": f"{passed}/{len(directories)} directories exist"
        }

    def _check_config(self):
        """Check configuration system"""
        logger.info("")
        logger.info("Checking Configuration System...")
        logger.info("-" * 60)

        passed = 0
        failed = 0
        total = 3

        try:
            from core.config import ConfigManager
            config = ConfigManager()
            logger.info(f"  ✓ ConfigManager initialized")
            passed += 1

            # Test reading config
            version = config.get('version', None)
            if version:
                logger.info(f"  ✓ Config readable (version: {version})")
                passed += 1
            else:
                logger.info(f"  ✗ Config readable: FAIL")
                failed += 1

            # Test AI config
            model = config.get('ai.model', None)
            if model:
                logger.info(f"  ✓ AI config present (model: {model})")
                passed += 1
            else:
                logger.info(f"  ✗ AI config present: FAIL")
                failed += 1

        except Exception as e:
            logger.info(f"  ✗ Configuration system: ERROR - {str(e)}")
            failed = total

        return {
            "status": "PASS" if failed == 0 else "FAIL",
            "checks_run": total,
            "checks_passed": passed,
            "details": f"{passed}/{total} config checks passed"
        }

    def _check_security(self):
        """Check security system"""
        logger.info("")
        logger.info("Checking Security System...")
        logger.info("-" * 60)

        passed = 0
        failed = 0
        total = 2

        try:
            from core.security import get_security_manager
            security = get_security_manager()
            logger.info(f"  ✓ SecurityManager initialized")
            passed += 1

            # Check if Windows Credential Manager is available
            try:
                import win32cred
                logger.info(f"  ✓ Windows Credential Manager available")
                passed += 1
            except ImportError:
                logger.info(f"  ! Windows Credential Manager: NOT AVAILABLE (using fallback)")
                # Not a failure, just informational
                passed += 1

        except Exception as e:
            logger.info(f"  ✗ Security system: ERROR - {str(e)}")
            failed = total

        return {
            "status": "PASS" if failed == 0 else "FAIL",
            "checks_run": total,
            "checks_passed": passed,
            "details": f"{passed}/{total} security checks passed"
        }

    def _check_modules(self):
        """Check automation modules"""
        logger.info("")
        logger.info("Checking Automation Modules...")
        logger.info("-" * 60)

        modules = [
            ('Desktop RPA', 'modules.desktop_rpa'),
            ('Excel Automation', 'modules.excel_automation'),
            ('Outlook Automation', 'modules.outlook_automation'),
            ('SharePoint', 'modules.sharepoint'),
            ('Word Automation', 'modules.word_automation'),
            ('OneNote', 'modules.onenote'),
            ('Base Module', 'modules.base_module')
        ]

        passed = 0
        failed = 0

        for name, module_path in modules:
            try:
                __import__(module_path)
                logger.info(f"  ✓ {name}: AVAILABLE")
                passed += 1
            except Exception as e:
                logger.info(f"  ✗ {name}: ERROR - {str(e)}")
                failed += 1

        return {
            "status": "PASS" if failed == 0 else "FAIL",
            "checks_run": len(modules),
            "checks_passed": passed,
            "details": f"{passed}/{len(modules)} modules available"
        }


def run(**kwargs):
    """Execute the system check"""
    test = SystemCheckTest()
    test.configure(**kwargs)
    test.validate()
    return test.execute()


if __name__ == "__main__":
    # Run the test
    result = run()
    print(f"\n\nTest Result: {result['status'].upper()}")
    print(f"Summary: {result['message']}")
