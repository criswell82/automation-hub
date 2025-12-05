"""
WORKFLOW_META:
  name: Script Discovery Test
  description: Validates the script discovery system can find and load custom workflows
  category: Tests
  version: 1.0.0
  author: Automation Hub
  parameters: {}
"""

import sys
from pathlib import Path

# Add src directory to path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class ScriptDiscoveryTest(BaseModule):
    """Test script discovery system"""

    def __init__(self):
        super().__init__()
        self.discovered_scripts = []

    def configure(self, **kwargs):
        """No configuration needed"""
        logger.info("Starting script discovery test...")
        return True

    def validate(self):
        """Always valid"""
        return True

    def execute(self):
        """Test script discovery"""
        try:
            logger.info("=" * 60)
            logger.info("SCRIPT DISCOVERY TEST")
            logger.info("=" * 60)

            # Test 1: Initialize discovery system
            logger.info("")
            logger.info("Test 1: Initialize Script Discovery")
            logger.info("-" * 60)

            test1_passed = False
            try:
                from core.script_discovery import ScriptDiscovery
                discovery = ScriptDiscovery()
                logger.info(f"✓ ScriptDiscovery initialized")
                logger.info(f"  Scripts directory: {discovery.scripts_dir}")
                test1_passed = True
            except Exception as e:
                logger.info(f"✗ ScriptDiscovery initialization failed: {e}")

            # Test 2: Scan for scripts
            logger.info("")
            logger.info("Test 2: Scan for Scripts")
            logger.info("-" * 60)

            test2_passed = False
            if test1_passed:
                try:
                    self.discovered_scripts = discovery.scan_directory()
                    logger.info(f"✓ Scan completed")
                    logger.info(f"  Found {len(self.discovered_scripts)} scripts")
                    test2_passed = True
                except Exception as e:
                    logger.info(f"✗ Scan failed: {e}")

            # Test 3: List discovered scripts
            logger.info("")
            logger.info("Test 3: List Discovered Scripts")
            logger.info("-" * 60)

            test3_passed = False
            if test2_passed and len(self.discovered_scripts) > 0:
                try:
                    # Group by category
                    by_category = {}
                    for script in self.discovered_scripts:
                        category = script.category
                        if category not in by_category:
                            by_category[category] = []
                        by_category[category].append(script)

                    logger.info(f"✓ Scripts organized by category")
                    logger.info("")

                    for category, scripts in sorted(by_category.items()):
                        logger.info(f"  {category}:")
                        for script in scripts:
                            logger.info(f"    • {script.name}")
                            logger.info(f"      ID: {script.id}")
                            logger.info(f"      File: {Path(script.file_path).name}")
                            logger.info(f"      Parameters: {len(script.parameters)}")
                            logger.info("")

                    test3_passed = True
                except Exception as e:
                    logger.info(f"✗ Listing failed: {e}")
            elif test2_passed:
                logger.info(f"! No scripts found to list")
                test3_passed = True  # Not a failure if no scripts exist yet

            # Test 4: Validate metadata
            logger.info("")
            logger.info("Test 4: Validate Script Metadata")
            logger.info("-" * 60)

            test4_passed = False
            if test2_passed:
                try:
                    valid_count = 0
                    invalid_count = 0

                    for script in self.discovered_scripts:
                        # Check required fields
                        has_id = bool(script.id)
                        has_name = bool(script.name)
                        has_category = bool(script.category)
                        has_file_path = bool(script.file_path)

                        if has_id and has_name and has_category and has_file_path:
                            valid_count += 1
                        else:
                            invalid_count += 1
                            logger.info(f"  ! Invalid metadata: {script.name}")
                            if not has_id:
                                logger.info(f"    - Missing ID")
                            if not has_name:
                                logger.info(f"    - Missing name")
                            if not has_category:
                                logger.info(f"    - Missing category")
                            if not has_file_path:
                                logger.info(f"    - Missing file path")

                    logger.info(f"✓ Metadata validation completed")
                    logger.info(f"  Valid: {valid_count}")
                    logger.info(f"  Invalid: {invalid_count}")

                    test4_passed = invalid_count == 0
                except Exception as e:
                    logger.info(f"✗ Validation failed: {e}")

            # Test 5: Test dynamic import
            logger.info("")
            logger.info("Test 5: Test Dynamic Import")
            logger.info("-" * 60)

            test5_passed = False
            if test2_passed and len(self.discovered_scripts) > 0:
                try:
                    # Try to import the first discovered script
                    test_script = self.discovered_scripts[0]
                    logger.info(f"  Attempting to load: {test_script.name}")

                    module = discovery.load_script_module(test_script)

                    if module:
                        logger.info(f"✓ Module loaded successfully")

                        # Check for run() function
                        if hasattr(module, 'run'):
                            logger.info(f"  ✓ run() function found")
                            test5_passed = True
                        else:
                            logger.info(f"  ! run() function not found")
                            test5_passed = True  # Still counts as successful load

                    else:
                        logger.info(f"✗ Module load returned None")

                except Exception as e:
                    logger.info(f"✗ Dynamic import failed: {e}")
                    import traceback
                    logger.info(f"  {traceback.format_exc()}")
            elif test2_passed:
                logger.info(f"! No scripts available to test import")
                test5_passed = True  # Not a failure

            # Summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("SUMMARY")
            logger.info("=" * 60)

            tests = [
                ("Initialize Discovery", test1_passed),
                ("Scan for Scripts", test2_passed),
                ("List Scripts", test3_passed),
                ("Validate Metadata", test4_passed),
                ("Dynamic Import", test5_passed)
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
                "message": f"Script discovery test: {passed}/{total} tests passed",
                "tests_passed": passed,
                "tests_total": total,
                "scripts_found": len(self.discovered_scripts),
                "test_results": {name: result for name, result in tests}
            }

        except Exception as e:
            self.handle_error(e, "Script discovery test failed")
            raise


def run(**kwargs):
    """Execute the test"""
    test = ScriptDiscoveryTest()
    test.configure(**kwargs)
    test.validate()
    return test.execute()


if __name__ == "__main__":
    # Run the test
    result = run()
    print(f"\n\nTest Result: {result['status'].upper()}")
    print(f"Summary: {result['message']}")
    print(f"Scripts Found: {result['scripts_found']}")
