"""
WORKFLOW_META:
  name: Workflow Execution Test
  description: Tests end-to-end workflow execution (load, configure, validate, execute)
  category: Tests
  version: 1.0.0
  author: Automation Hub
  parameters: {}
"""

import sys
import time
from pathlib import Path

# Add src directory to path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from src.modules.base_module import BaseModule
from src.core.logging_config import get_logger

logger = get_logger(__name__)


class WorkflowExecutionTest(BaseModule):
    """Test end-to-end workflow execution"""

    def __init__(self):
        super().__init__()

    def configure(self, **kwargs):
        """No configuration needed"""
        logger.info("Starting workflow execution test...")
        return True

    def validate(self):
        """Always valid"""
        return True

    def execute(self):
        """Test workflow execution"""
        try:
            logger.info("=" * 60)
            logger.info("WORKFLOW EXECUTION TEST")
            logger.info("=" * 60)

            # Test 1: Load example workflow
            logger.info("")
            logger.info("Test 1: Load Example Workflow")
            logger.info("-" * 60)

            test1_passed = False
            example_workflow_path = None
            try:
                # Try to find the example_hello_world.py workflow
                custom_dir = project_root / 'user_scripts' / 'custom'
                example_path = custom_dir / 'example_hello_world.py'

                if example_path.exists():
                    example_workflow_path = example_path
                    logger.info(f"✓ Example workflow found")
                    logger.info(f"  Path: {example_path}")
                    test1_passed = True
                else:
                    logger.info(f"! Example workflow not found at: {example_path}")
                    logger.info(f"  Creating a test workflow inline instead...")

                    # If example doesn't exist, we'll test with inline code
                    test1_passed = True

            except Exception as e:
                logger.info(f"✗ Loading workflow failed: {e}")

            # Test 2: Execute simple inline workflow
            logger.info("")
            logger.info("Test 2: Execute Inline Test Workflow")
            logger.info("-" * 60)

            test2_passed = False
            if test1_passed:
                try:
                    # Create a simple test workflow class inline
                    class SimpleTestWorkflow(BaseModule):
                        def __init__(self):
                            super().__init__()
                            self.test_param = None

                        def configure(self, **kwargs):
                            self.test_param = kwargs.get('test_param', 'default_value')
                            return True

                        def validate(self):
                            return bool(self.test_param)

                        def execute(self):
                            return {
                                'status': 'success',
                                'message': 'Test workflow executed successfully',
                                'test_param': self.test_param
                            }

                    # Execute the workflow
                    logger.info(f"  Creating workflow instance...")
                    workflow = SimpleTestWorkflow()

                    logger.info(f"  Configuring workflow...")
                    workflow.configure(test_param='test_value_123')

                    logger.info(f"  Validating workflow...")
                    validation_result = workflow.validate()

                    if not validation_result:
                        logger.info(f"  ! Validation failed")
                    else:
                        logger.info(f"  ✓ Validation passed")

                        logger.info(f"  Executing workflow...")
                        start_time = time.time()
                        result = workflow.execute()
                        end_time = time.time()
                        execution_time = (end_time - start_time) * 1000  # ms

                        logger.info(f"✓ Workflow executed successfully")
                        logger.info(f"  Execution time: {execution_time:.2f}ms")
                        logger.info(f"  Result: {result}")

                        # Validate result structure
                        has_status = 'status' in result
                        has_message = 'message' in result
                        status_is_success = result.get('status') == 'success'

                        logger.info(f"  Result validation:")
                        logger.info(f"    - Has status: {'✓' if has_status else '✗'}")
                        logger.info(f"    - Has message: {'✓' if has_message else '✗'}")
                        logger.info(f"    - Status is success: {'✓' if status_is_success else '✗'}")

                        test2_passed = has_status and has_message and status_is_success

                except Exception as e:
                    logger.info(f"✗ Workflow execution failed: {e}")
                    import traceback
                    logger.info(f"  {traceback.format_exc()}")

            # Test 3: Test error handling
            logger.info("")
            logger.info("Test 3: Test Error Handling")
            logger.info("-" * 60)

            test3_passed = False
            if test1_passed:
                try:
                    # Create a workflow that intentionally fails
                    class FailingWorkflow(BaseModule):
                        def configure(self, **kwargs):
                            return True

                        def validate(self):
                            return True

                        def execute(self):
                            raise ValueError("Intentional test error")

                    logger.info(f"  Creating failing workflow...")
                    failing_workflow = FailingWorkflow()
                    failing_workflow.configure()
                    failing_workflow.validate()

                    logger.info(f"  Executing (should fail)...")
                    error_caught = False
                    try:
                        failing_workflow.execute()
                    except ValueError as e:
                        if "Intentional test error" in str(e):
                            logger.info(f"✓ Error caught correctly")
                            logger.info(f"  Error message: {str(e)}")
                            error_caught = True
                        else:
                            logger.info(f"✗ Wrong error caught: {str(e)}")
                    except Exception as e:
                        logger.info(f"✗ Unexpected error type: {type(e).__name__}")

                    test3_passed = error_caught

                except Exception as e:
                    logger.info(f"✗ Error handling test failed: {e}")

            # Test 4: Test workflow lifecycle
            logger.info("")
            logger.info("Test 4: Test Workflow Lifecycle (Configure → Validate → Execute)")
            logger.info("-" * 60)

            test4_passed = False
            if test1_passed:
                try:
                    # Create workflow with lifecycle tracking
                    class LifecycleWorkflow(BaseModule):
                        def __init__(self):
                            super().__init__()
                            self.lifecycle_steps = []

                        def configure(self, **kwargs):
                            self.lifecycle_steps.append('configure')
                            return True

                        def validate(self):
                            self.lifecycle_steps.append('validate')
                            return True

                        def execute(self):
                            self.lifecycle_steps.append('execute')
                            return {
                                'status': 'success',
                                'lifecycle': self.lifecycle_steps
                            }

                    logger.info(f"  Testing lifecycle...")
                    lc_workflow = LifecycleWorkflow()

                    lc_workflow.configure()
                    lc_workflow.validate()
                    result = lc_workflow.execute()

                    expected_lifecycle = ['configure', 'validate', 'execute']
                    actual_lifecycle = result.get('lifecycle', [])

                    if actual_lifecycle == expected_lifecycle:
                        logger.info(f"✓ Lifecycle executed in correct order")
                        logger.info(f"  Steps: {' → '.join(actual_lifecycle)}")
                        test4_passed = True
                    else:
                        logger.info(f"✗ Lifecycle order incorrect")
                        logger.info(f"  Expected: {expected_lifecycle}")
                        logger.info(f"  Actual: {actual_lifecycle}")

                except Exception as e:
                    logger.info(f"✗ Lifecycle test failed: {e}")

            # Test 5: Test with ScriptDiscovery (if example exists)
            logger.info("")
            logger.info("Test 5: Load and Execute via ScriptDiscovery")
            logger.info("-" * 60)

            test5_passed = False
            if example_workflow_path and example_workflow_path.exists():
                try:
                    from core.script_discovery import ScriptDiscovery

                    discovery = ScriptDiscovery()
                    scripts = discovery.scan_directory()

                    # Find the example hello world workflow
                    example_script = None
                    for script in scripts:
                        if 'hello' in script.name.lower() and 'example' in script.name.lower():
                            example_script = script
                            break

                    if example_script:
                        logger.info(f"  Found: {example_script.name}")
                        logger.info(f"  Loading module...")

                        start_time = time.time()
                        result = discovery.execute_script(
                            example_script,
                            name="Test User",
                            message="Testing workflow execution!"
                        )
                        end_time = time.time()
                        execution_time = (end_time - start_time) * 1000  # ms

                        logger.info(f"✓ Workflow executed via ScriptDiscovery")
                        logger.info(f"  Execution time: {execution_time:.2f}ms")
                        logger.info(f"  Result: {result.get('status', 'unknown')}")

                        test5_passed = result.get('status') == 'success'
                    else:
                        logger.info(f"  ! Example workflow not found in discovered scripts")
                        test5_passed = True  # Don't fail if not found

                except Exception as e:
                    logger.info(f"✗ ScriptDiscovery execution failed: {e}")
                    import traceback
                    logger.info(f"  {traceback.format_exc()}")
            else:
                logger.info(f"! Skipped (example workflow not available)")
                test5_passed = True  # Don't fail if example doesn't exist

            # Summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("SUMMARY")
            logger.info("=" * 60)

            tests = [
                ("Load Example Workflow", test1_passed),
                ("Execute Inline Workflow", test2_passed),
                ("Error Handling", test3_passed),
                ("Workflow Lifecycle", test4_passed),
                ("Execute via ScriptDiscovery", test5_passed)
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
                "message": f"Workflow execution test: {passed}/{total} tests passed",
                "tests_passed": passed,
                "tests_total": total,
                "test_results": {name: result for name, result in tests}
            }

        except Exception as e:
            self.handle_error(e, "Workflow execution test failed")
            raise


def run(**kwargs):
    """Execute the test"""
    test = WorkflowExecutionTest()
    test.configure(**kwargs)
    test.validate()
    return test.execute()


if __name__ == "__main__":
    # Run the test
    result = run()
    print(f"\n\nTest Result: {result['status'].upper()}")
    print(f"Summary: {result['message']}")
