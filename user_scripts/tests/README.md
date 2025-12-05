# Automation Hub Test Workflows

This directory contains validation test workflows that verify your Automation Hub installation is working correctly.

## Available Tests

### 1. System Check (`system_check.py`)
**Purpose**: Validates your entire system setup

**What it tests:**
- All required Python packages are installed
- Directory structure is correct (config, logs, user_scripts, etc.)
- Configuration system works
- Security system (Windows Credential Manager) is available
- All automation modules can be imported

**When to run**: After initial installation or when troubleshooting setup issues

---

### 2. Script Discovery Test (`script_discovery_test.py`)
**Purpose**: Tests the script discovery system

**What it tests:**
- Script scanning works correctly
- YAML metadata parsing from docstrings
- Script categorization
- Dynamic module loading
- Lists all discovered workflows

**When to run**: After adding custom workflows or when workflows don't appear in the GUI

---

### 3. AI Generator Test (`ai_generator_test.py`)
**Purpose**: Tests AI workflow generation

**What it tests:**
- AI Workflow Generator initialization
- API key detection
- Available models list
- Template-based generation (always works, no API key needed)
- AI-powered generation (if API key configured)
- Generated code validation

**When to run**: After configuring your API key or when workflow generation fails

**Note**: Some tests will be skipped if no API key is configured (this is expected)

---

### 4. Configuration Test (`config_test.py`)
**Purpose**: Tests the configuration management system

**What it tests:**
- ConfigManager initialization
- Reading default configuration values
- Writing and reading custom values
- Dot-notation access (e.g., 'ai.model')
- Config file locations and permissions
- Module-specific configuration

**When to run**: When settings don't save or config values seem incorrect

---

### 5. Workflow Execution Test (`workflow_execution_test.py`)
**Purpose**: Tests end-to-end workflow execution

**What it tests:**
- Loading workflows
- Workflow lifecycle (configure → validate → execute)
- Result structure validation
- Error handling
- Execution via ScriptDiscovery
- Execution timing

**When to run**: When workflows fail to execute or produce unexpected results

---

## How to Run Tests

### From the GUI (Recommended)

1. **Launch Automation Hub**
   ```bash
   python src/main.py
   ```

2. **Find the Tests Category**
   - Look in the script library on the left
   - Expand the "Tests" category

3. **Select a Test**
   - Click on any test workflow (e.g., "System Check Test")

4. **Run the Test**
   - Click the "Run Script" button (or press Ctrl+R)
   - Watch the output in the Output tab
   - Look for ✓ PASS or ✗ FAIL indicators

5. **Review Results**
   - Each test provides a summary at the end
   - Clear pass/fail status for each check
   - Helpful error messages if something fails

### From Command Line

You can also run tests directly with Python:

```bash
# System Check
python user_scripts/tests/system_check.py

# Script Discovery
python user_scripts/tests/script_discovery_test.py

# AI Generator
python user_scripts/tests/ai_generator_test.py

# Configuration
python user_scripts/tests/config_test.py

# Workflow Execution
python user_scripts/tests/workflow_execution_test.py
```

---

## Interpreting Results

### Test Output Symbols

- **✓** - Test passed
- **✗** - Test failed
- **!** - Informational/Warning (not a failure)

### Status Codes

- **SUCCESS** - All tests passed
- **WARNING** - Some tests failed, but non-critical
- **FAIL** - Critical tests failed

### Example Output

```
==========================================================
AUTOMATION HUB SYSTEM CHECK
==========================================================

Checking Dependencies...
----------------------------------------------------------
  ✓ PyQt5: INSTALLED
  ✓ anthropic: INSTALLED
  ✓ openpyxl: INSTALLED
  ...

==========================================================
SUMMARY
==========================================================
Total Checks: 23
✓ Passed: 23
✗ Failed: 0
==========================================================
```

---

## Common Issues and Solutions

### All Dependencies Test Fails
**Problem**: Some packages are missing

**Solution**:
```bash
pip install -r requirements.txt
```

### API Key Tests Skipped
**Problem**: No API key configured (this is okay!)

**Solution**:
- Open Settings (Ctrl+,)
- Go to API Keys tab
- Enter your Anthropic API key
- Run test again

### Config Directory Tests Fail
**Problem**: Config directory doesn't exist or isn't writable

**Solution**:
- Check permissions on `%APPDATA%/AutomationHub`
- Run as administrator if needed
- Check disk space

### Script Discovery Finds No Scripts
**Problem**: No custom workflows exist yet

**Solution**: This is normal for new installations. The templates should still be found.

### Workflow Execution Fails
**Problem**: Example workflow missing

**Solution**: This is normal if you haven't created any custom workflows yet. The inline tests should still pass.

---

## Test Run Recommendations

### First Time Setup
Run tests in this order:
1. System Check - Verify everything is installed
2. Configuration Test - Ensure config system works
3. Script Discovery Test - Verify workflow loading works
4. Workflow Execution Test - Test running workflows
5. AI Generator Test - Test AI features (optional)

### After Adding API Key
Run:
- AI Generator Test

### After Creating Custom Workflows
Run:
- Script Discovery Test
- Workflow Execution Test

### Before Reporting Issues
Run ALL tests and include the results when reporting problems.

---

## What Each Test Category Means

### Tests Category in GUI
All test workflows appear under the "Tests" category in the script library. This makes them easy to find and run.

### No Parameters Required
None of the test workflows require parameters - just click Run!

### Safe to Run Multiple Times
All tests are non-destructive and safe to run repeatedly. They don't modify your data or settings (except for temporary test values that are cleaned up).

---

## Advanced: Running All Tests at Once

You can run all tests sequentially from command line:

```bash
cd automation_hub

# Run all tests
for test in user_scripts/tests/*.py; do
    echo "Running $test..."
    python "$test"
    echo ""
done
```

Or create a simple batch file (`run_all_tests.bat`):

```batch
@echo off
echo Running All Validation Tests...
echo.

python user_scripts/tests/system_check.py
python user_scripts/tests/script_discovery_test.py
python user_scripts/tests/config_test.py
python user_scripts/tests/workflow_execution_test.py
python user_scripts/tests/ai_generator_test.py

echo.
echo All tests completed!
pause
```

---

## Getting Help

If tests fail and you can't figure out why:

1. **Check the logs** - Detailed error messages are in `%APPDATA%/AutomationHub/logs`
2. **Review test output** - Look for specific error messages
3. **Check requirements** - Make sure all dependencies are installed
4. **Try command line** - Sometimes running from command line shows more details

---

## For Developers

These test workflows are also great examples of:
- How to structure workflows
- How to use BaseModule
- How to add logging
- How to validate and handle errors
- How to work with the Automation Hub APIs

Feel free to copy and modify them for your own testing needs!

---

**Happy Testing! 🧪**
