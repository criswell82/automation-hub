"""
Build script for creating Automation Hub executable.
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Build the application using PyInstaller."""
    print("=" * 60)
    print("Automation Hub - Build Script")
    print("=" * 60)

    # Get project root
    script_dir = Path(__file__).parent
    project_root = script_dir.parent

    # Change to project root
    os.chdir(project_root)

    print(f"\nProject root: {project_root}")
    print(f"Python: {sys.executable}")
    print(f"Python version: {sys.version}")

    # Clean previous builds
    print("\n[1/4] Cleaning previous builds...")
    for dir_name in ['build', 'dist']:
        dir_path = project_root / dir_name
        if dir_path.exists():
            print(f"  Removing {dir_path}")
            import shutil
            shutil.rmtree(dir_path)

    # Install/upgrade dependencies
    print("\n[2/4] Checking dependencies...")
    subprocess.run([
        sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "--upgrade"
    ], check=True)

    # Run PyInstaller
    print("\n[3/4] Running PyInstaller...")
    spec_file = project_root / "pyinstaller.spec"

    if not spec_file.exists():
        print(f"ERROR: Spec file not found: {spec_file}")
        return 1

    subprocess.run([
        sys.executable, "-m", "PyInstaller",
        str(spec_file),
        "--clean",
        "--noconfirm"
    ], check=True)

    # Verify output
    print("\n[4/4] Verifying build...")
    exe_path = project_root / "dist" / "AutomationHub.exe"

    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\n[OK] Build successful!")
        print(f"  Executable: {exe_path}")
        print(f"  Size: {size_mb:.2f} MB")
        return 0
    else:
        print(f"\n[FAIL] Build failed - executable not found!")
        return 1

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except Exception as e:
        print(f"\n[FAIL] Build failed with error: {e}")
        sys.exit(1)
