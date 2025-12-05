"""
WORKFLOW_META:
  name: AI Generator Test
  description: Tests AI workflow generation system (template-based and AI-powered)
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


class AIGeneratorTest(BaseModule):
    """Test AI workflow generation system"""

    def __init__(self):
        super().__init__()

    def configure(self, **kwargs):
        """No configuration needed"""
        logger.info("Starting AI generator test...")
        return True

    def validate(self):
        """Always valid"""
        return True

    def execute(self):
        """Test AI workflow generator"""
        try:
            logger.info("=" * 60)
            logger.info("AI WORKFLOW GENERATOR TEST")
            logger.info("=" * 60)

            # Test 1: Initialize generator
            logger.info("")
            logger.info("Test 1: Initialize AI Workflow Generator")
            logger.info("-" * 60)

            test1_passed = False
            generator = None
            try:
                from core.ai_workflow_generator import AIWorkflowGenerator
                generator = AIWorkflowGenerator()
                logger.info(f"✓ AIWorkflowGenerator initialized")
                test1_passed = True
            except Exception as e:
                logger.info(f"✗ Initialization failed: {e}")

            # Test 2: Check API key availability
            logger.info("")
            logger.info("Test 2: Check API Key Availability")
            logger.info("-" * 60)

            has_api_key = False
            if test1_passed:
                try:
                    has_api_key = bool(generator.api_key)
                    if has_api_key:
                        logger.info(f"✓ API key is configured")
                        logger.info(f"  Key starts with: {generator.api_key[:10]}...")
                    else:
                        logger.info(f"! No API key configured (template mode will be used)")
                except Exception as e:
                    logger.info(f"✗ API key check failed: {e}")

            # Test 3: Get available models
            logger.info("")
            logger.info("Test 3: Get Available Models")
            logger.info("-" * 60)

            test3_passed = False
            if test1_passed:
                try:
                    models = AIWorkflowGenerator.get_available_models()
                    logger.info(f"✓ Retrieved {len(models)} models")
                    logger.info("")
                    for model in models:
                        logger.info(f"  • {model['name']}")
                        logger.info(f"    ID: {model['id']}")
                        logger.info(f"    Description: {model['description']}")
                        logger.info("")
                    test3_passed = True
                except Exception as e:
                    logger.info(f"✗ Getting models failed: {e}")

            # Test 4: Template-based generation (always works)
            logger.info("")
            logger.info("Test 4: Template-Based Generation (No API Key)")
            logger.info("-" * 60)

            test4_passed = False
            template_code = None
            if test1_passed:
                try:
                    test_description = "Create a simple hello world workflow"
                    logger.info(f"  Description: {test_description}")

                    # Force template generation by using _generate_from_template
                    result = generator._generate_from_template(test_description, "Custom")

                    if result and 'code' in result:
                        template_code = result['code']
                        logger.info(f"✓ Template generation successful")
                        logger.info(f"  Generated code length: {len(template_code)} characters")
                        logger.info(f"  Workflow name: {result.get('name', 'N/A')}")

                        # Validate generated code
                        has_workflow_meta = 'WORKFLOW_META' in template_code
                        has_base_module = 'BaseModule' in template_code
                        has_run_function = 'def run(' in template_code

                        logger.info(f"  Validation:")
                        logger.info(f"    - Has WORKFLOW_META: {'✓' if has_workflow_meta else '✗'}")
                        logger.info(f"    - Has BaseModule: {'✓' if has_base_module else '✗'}")
                        logger.info(f"    - Has run() function: {'✓' if has_run_function else '✗'}")

                        test4_passed = has_workflow_meta and has_base_module and has_run_function
                    else:
                        logger.info(f"✗ Template generation returned invalid result")

                except Exception as e:
                    logger.info(f"✗ Template generation failed: {e}")
                    import traceback
                    logger.info(f"  {traceback.format_exc()}")

            # Test 5: AI-powered generation (only if API key available)
            logger.info("")
            logger.info("Test 5: AI-Powered Generation (With API Key)")
            logger.info("-" * 60)

            test5_passed = False
            test5_skipped = False

            if test1_passed and has_api_key:
                try:
                    test_description = "Create a workflow that prints hello world"
                    logger.info(f"  Description: {test_description}")
                    logger.info(f"  Calling Claude API...")

                    result = generator.generate_workflow(
                        description=test_description,
                        category="Custom",
                        use_templates=[]
                    )

                    if result and 'code' in result:
                        ai_code = result['code']
                        logger.info(f"✓ AI generation successful")
                        logger.info(f"  Generated code length: {len(ai_code)} characters")
                        logger.info(f"  Workflow name: {result.get('name', 'N/A')}")

                        # Validate
                        has_workflow_meta = 'WORKFLOW_META' in ai_code
                        has_base_module = 'BaseModule' in ai_code
                        has_run_function = 'def run(' in ai_code

                        logger.info(f"  Validation:")
                        logger.info(f"    - Has WORKFLOW_META: {'✓' if has_workflow_meta else '✗'}")
                        logger.info(f"    - Has BaseModule: {'✓' if has_base_module else '✗'}")
                        logger.info(f"    - Has run() function: {'✓' if has_run_function else '✗'}")

                        test5_passed = has_workflow_meta and has_base_module and has_run_function
                    else:
                        logger.info(f"✗ AI generation returned invalid result")

                except Exception as e:
                    logger.info(f"✗ AI generation failed: {e}")
                    logger.info(f"  Error details: {str(e)}")
            else:
                logger.info(f"! Skipped (no API key configured)")
                logger.info(f"  To test AI generation, add your API key in Settings")
                test5_skipped = True
                test5_passed = True  # Don't count as failure

            # Summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("SUMMARY")
            logger.info("=" * 60)

            tests = [
                ("Initialize Generator", test1_passed, False),
                ("API Key Check", True, False),  # Always pass, just informational
                ("Get Available Models", test3_passed, False),
                ("Template Generation", test4_passed, False),
                ("AI Generation", test5_passed, test5_skipped)
            ]

            passed = sum(1 for _, result, skipped in tests if result and not skipped)
            skipped = sum(1 for _, result, is_skipped in tests if is_skipped)
            failed = len(tests) - passed - skipped

            for test_name, result, is_skipped in tests:
                if is_skipped:
                    logger.info(f"  ! SKIP: {test_name}")
                elif result:
                    logger.info(f"  ✓ PASS: {test_name}")
                else:
                    logger.info(f"  ✗ FAIL: {test_name}")

            logger.info("")
            logger.info(f"Passed: {passed}")
            logger.info(f"Failed: {failed}")
            logger.info(f"Skipped: {skipped}")
            logger.info(f"API Key: {'Configured' if has_api_key else 'Not Configured'}")
            logger.info("=" * 60)

            # Consider it success if all non-skipped tests passed
            all_passed = failed == 0

            return {
                "status": "success" if all_passed else "warning",
                "message": f"AI generator test: {passed}/{len(tests)-skipped} tests passed",
                "tests_passed": passed,
                "tests_failed": failed,
                "tests_skipped": skipped,
                "tests_total": len(tests),
                "has_api_key": has_api_key,
                "template_code_length": len(template_code) if template_code else 0
            }

        except Exception as e:
            self.handle_error(e, "AI generator test failed")
            raise


def run(**kwargs):
    """Execute the test"""
    test = AIGeneratorTest()
    test.configure(**kwargs)
    test.validate()
    return test.execute()


if __name__ == "__main__":
    # Run the test
    result = run()
    print(f"\n\nTest Result: {result['status'].upper()}")
    print(f"Summary: {result['message']}")
