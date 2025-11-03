"""
Example: Copy-Paste Automation
Demonstrates copying data from one application and pasting into another.
This is a common use case for corporate environments with locked-down tools.
"""

import sys
from pathlib import Path
import time

# Add parent directories to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

from modules.desktop_rpa import WindowManager, InputController


def copy_paste_automation():
    """
    Example: Demonstrate copy-paste workflow between applications.

    This example shows how to:
    1. Find a source window
    2. Select and copy text
    3. Switch to target window
    4. Paste the text
    """
    print("=== Copy-Paste Automation Example ===\n")

    wm = WindowManager(default_timeout=10)
    ic = InputController(default_delay=0.5)

    try:
        # Step 1: Get the active window (source)
        print("1. Please focus on the SOURCE window (where you want to copy from)")
        print("   Waiting 5 seconds...")
        time.sleep(5)

        source_window = wm.get_active_window()
        if not source_window:
            print("ERROR: Could not get source window!")
            return

        print(f"   Source window: {source_window.title}")

        # Step 2: Select all and copy
        print("\n2. Selecting all text and copying...")
        ic.hotkey('ctrl', 'a')  # Select all
        time.sleep(0.3)
        ic.hotkey('ctrl', 'c')  # Copy
        time.sleep(0.3)

        print("   Text copied to clipboard!")

        # Step 3: Switch to target window
        print("\n3. Please focus on the TARGET window (where you want to paste)")
        print("   Waiting 5 seconds...")
        time.sleep(5)

        target_window = wm.get_active_window()
        if not target_window:
            print("ERROR: Could not get target window!")
            return

        print(f"   Target window: {target_window.title}")

        # Step 4: Paste
        print("\n4. Pasting text...")
        ic.hotkey('ctrl', 'v')  # Paste
        time.sleep(0.5)

        print("\n5. Automation complete!")
        print("   Text has been copied from source to target.")

    except Exception as e:
        print(f"ERROR: {e}")


def automated_data_entry_example():
    """
    Example: Automated data entry into a form.
    Simulates filling out multiple fields using Tab navigation.
    """
    print("\n=== Automated Data Entry Example ===\n")

    ic = InputController(default_delay=0.3)

    # Sample data to enter
    data_fields = [
        "John Doe",                      # Name
        "john.doe@company.com",          # Email
        "555-1234",                      # Phone
        "123 Main Street",               # Address
        "New York",                      # City
        "10001",                         # ZIP
    ]

    try:
        print("Position cursor in the FIRST FIELD of your form...")
        print("Starting in 5 seconds...")
        time.sleep(5)

        print("\nEntering data...")
        for i, value in enumerate(data_fields, 1):
            print(f"  Field {i}: {value}")
            ic.type_text(value, interval=0.05)
            ic.press_key('tab')  # Move to next field
            time.sleep(0.2)

        print("\nData entry complete!")

    except Exception as e:
        print(f"ERROR: {e}")


if __name__ == "__main__":
    print("Desktop RPA Copy-Paste Examples\n")
    print("Choose an example:")
    print("1. Copy-Paste between windows")
    print("2. Automated data entry")

    choice = input("\nEnter choice (1 or 2): ").strip()

    if choice == "1":
        print("\n" + "=" * 50)
        copy_paste_automation()
    elif choice == "2":
        print("\n" + "=" * 50)
        automated_data_entry_example()
    else:
        print("Invalid choice!")
