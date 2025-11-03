"""
Example: Automating Notepad
Demonstrates basic window management and text input.
"""

import sys
from pathlib import Path
import subprocess
import time

# Add parent directories to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

from modules.desktop_rpa import WindowManager, InputController


def automate_notepad():
    """
    Example automation: Launch Notepad, type text, and save.
    """
    print("=== Notepad Automation Example ===\n")

    # Initialize controllers
    wm = WindowManager(default_timeout=10)
    ic = InputController(default_delay=0.5)

    try:
        # Step 1: Launch Notepad
        print("1. Launching Notepad...")
        subprocess.Popen(['notepad.exe'])

        # Step 2: Wait for Notepad window to appear
        print("2. Waiting for Notepad window...")
        window = wm.wait_for_window(title_regex=".*Notepad.*", timeout=10)

        if not window:
            print("ERROR: Notepad window not found!")
            return

        print(f"   Found window: {window.title}")

        # Step 3: Activate the window
        print("3. Activating window...")
        wm.activate_window(window)
        time.sleep(0.5)

        # Step 4: Type some text
        print("4. Typing text...")
        sample_text = """Automation Hub - Desktop RPA Example

This text was automatically typed using the Desktop RPA module!

Features demonstrated:
- Window management
- Text input automation
- Keyboard shortcuts

Timestamp: """ + time.strftime("%Y-%m-%d %H:%M:%S")

        ic.type_text(sample_text, interval=0.01)

        print("5. Automation complete!")
        print("\nNOTE: Notepad window left open for inspection.")
        print("      Close it manually or press Ctrl+W")

    except Exception as e:
        print(f"ERROR: {e}")


if __name__ == "__main__":
    print("Starting automation in 3 seconds...")
    print("(Move mouse to top-left corner to abort)")
    time.sleep(3)

    automate_notepad()
