"""
Example: Advanced RPA Automation
Demonstrates advanced features like window positioning, screenshots, and retry logic.
"""

import sys
from pathlib import Path
import time

# Add parent directories to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

from modules.desktop_rpa import WindowManager, InputController


def window_management_example():
    """
    Demonstrate advanced window management features.
    """
    print("=== Window Management Example ===\n")

    wm = WindowManager(default_timeout=10)

    try:
        # List all windows
        print("1. Listing all visible windows...")
        windows = wm.get_all_windows(visible_only=True)

        print(f"\nFound {len(windows)} windows:")
        for i, window in enumerate(windows[:10], 1):  # Show first 10
            print(f"   {i}. {window.title[:50]} (Class: {window.class_name})")

        if len(windows) > 10:
            print(f"   ... and {len(windows) - 10} more")

        # Find a specific window by regex
        print("\n2. Searching for a Chrome or Edge window...")
        browser_window = wm.find_window(title_regex=".*(Chrome|Edge).*", timeout=5)

        if browser_window:
            print(f"   Found: {browser_window.title}")

            # Get window position
            rect = wm.get_window_rect(browser_window)
            if rect:
                left, top, right, bottom = rect
                width = right - left
                height = bottom - top
                print(f"   Position: ({left}, {top})")
                print(f"   Size: {width}x{height}")

            # Demonstrate window operations
            print("\n3. Demonstrating window operations...")
            print("   - Activating window...")
            wm.activate_window(browser_window)
            time.sleep(1)

            print("   - Moving window to (100, 100)...")
            wm.set_window_position(100, 100, 800, 600, browser_window)
            time.sleep(1)

            print("   - Maximizing window...")
            wm.maximize_window(browser_window)
            time.sleep(1)

        else:
            print("   No browser window found.")

    except Exception as e:
        print(f"ERROR: {e}")


def screenshot_and_locate_example():
    """
    Demonstrate screenshot and image location features.
    """
    print("\n=== Screenshot Example ===\n")

    ic = InputController()

    try:
        # Take a screenshot
        print("1. Taking full screenshot...")
        screenshot_path = ic.screenshot()
        if screenshot_path:
            print(f"   Saved to: {screenshot_path}")

        # Take a region screenshot
        print("\n2. Taking region screenshot...")
        # Capture center of screen (example: 500x500 region)
        x, y = ic.get_mouse_position()
        region = (x - 250, y - 250, 500, 500)
        region_screenshot = ic.screenshot(
            filepath="region_screenshot.png",
            region=region
        )
        if region_screenshot:
            print(f"   Saved to: {region_screenshot}")

        print("\nScreenshot examples complete!")

    except Exception as e:
        print(f"ERROR: {e}")


def retry_logic_example():
    """
    Demonstrate retry logic for unreliable operations.
    """
    print("\n=== Retry Logic Example ===\n")

    wm = WindowManager(default_timeout=5)
    ic = InputController(default_delay=0.3)

    def try_click_button(x, y, max_retries=3):
        """
        Try to click a button with retry logic.
        """
        for attempt in range(1, max_retries + 1):
            print(f"   Attempt {attempt}/{max_retries}...")

            # Move to position
            ic.move_to(x, y)
            time.sleep(0.2)

            # Take a screenshot to verify
            screenshot = ic.screenshot(f"click_attempt_{attempt}.png")

            # Click
            if ic.click(x, y):
                print(f"   Click successful on attempt {attempt}")
                return True

            # Wait before retry
            if attempt < max_retries:
                print(f"   Retrying in 1 second...")
                time.sleep(1)

        print("   All attempts failed!")
        return False

    try:
        print("This example shows how to implement retry logic.")
        print("Note: This is a simulation - customize coordinates for your use case.\n")

        # Example: Try to click at a specific position
        target_x, target_y = 500, 300
        print(f"Attempting to click at ({target_x}, {target_y}) with retry logic...")

        success = try_click_button(target_x, target_y, max_retries=3)

        if success:
            print("\nOperation completed successfully!")
        else:
            print("\nOperation failed after all retries.")

    except Exception as e:
        print(f"ERROR: {e}")


def safe_automation_pattern():
    """
    Demonstrate a safe automation pattern with error handling.
    """
    print("\n=== Safe Automation Pattern ===\n")

    wm = WindowManager()
    ic = InputController()

    try:
        print("This demonstrates a robust automation pattern:\n")

        # 1. Verify window exists
        print("1. Verify target window exists...")
        window = wm.get_active_window()
        if not window:
            raise Exception("No active window found!")
        print(f"   ✓ Window found: {window.title}")

        # 2. Take a before screenshot
        print("\n2. Capture initial state...")
        before_screenshot = ic.screenshot("before.png")
        print(f"   ✓ Screenshot saved: {before_screenshot}")

        # 3. Activate window
        print("\n3. Activate and focus window...")
        if not wm.activate_window(window):
            raise Exception("Failed to activate window!")
        print("   ✓ Window activated")

        # 4. Perform operation with verification
        print("\n4. Perform operation...")
        ic.wait(0.5)
        print("   ✓ Operation simulated")

        # 5. Take an after screenshot
        print("\n5. Capture final state...")
        after_screenshot = ic.screenshot("after.png")
        print(f"   ✓ Screenshot saved: {after_screenshot}")

        # 6. Verify success (in real scenario, check window state or content)
        print("\n6. Verify operation success...")
        print("   ✓ Verification passed")

        print("\n✓ Automation completed successfully with full error handling!")

    except Exception as e:
        print(f"\n✗ Automation failed: {e}")
        print("   Taking error screenshot...")
        ic.screenshot("error.png")


if __name__ == "__main__":
    print("Advanced Desktop RPA Examples\n")
    print("=" * 60)

    examples = {
        "1": ("Window Management", window_management_example),
        "2": ("Screenshots", screenshot_and_locate_example),
        "3": ("Retry Logic", retry_logic_example),
        "4": ("Safe Automation Pattern", safe_automation_pattern),
        "5": ("Run All Examples", None)
    }

    print("\nAvailable examples:")
    for key, (name, _) in examples.items():
        print(f"  {key}. {name}")

    choice = input("\nEnter choice (1-5): ").strip()

    print("\n" + "=" * 60 + "\n")

    if choice == "5":
        # Run all examples
        for key in ["1", "2", "3", "4"]:
            examples[key][1]()
            print("\n" + "=" * 60 + "\n")
            time.sleep(2)
    elif choice in examples and examples[choice][1]:
        examples[choice][1]()
    else:
        print("Invalid choice!")
