"""
Window Manager for Desktop RPA.
Handles window detection, focus, positioning, and management using pywin32.
"""

import logging
import time
from typing import Optional, List, Tuple, Dict, Any
import re

try:
    import win32gui
    import win32con
    import win32process
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    logging.warning("pywin32 not available - window management features will be limited")


class WindowInfo:
    """Represents information about a window."""

    def __init__(self, hwnd: int, title: str, class_name: str, process_id: int):
        """
        Initialize window information.

        Args:
            hwnd: Window handle
            title: Window title
            class_name: Window class name
            process_id: Process ID
        """
        self.hwnd = hwnd
        self.title = title
        self.class_name = class_name
        self.process_id = process_id

    def __str__(self) -> str:
        return f"Window(hwnd={self.hwnd}, title='{self.title}', class='{self.class_name}')"

    def __repr__(self) -> str:
        return self.__str__()


class WindowManager:
    """
    Manages application windows for RPA automation.
    Provides methods to find, activate, position, and interact with windows.
    """

    def __init__(self, default_timeout: int = 10):
        """
        Initialize the window manager.

        Args:
            default_timeout: Default timeout for window operations in seconds
        """
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 is required for WindowManager")

        self.logger = logging.getLogger(__name__)
        self.default_timeout = default_timeout
        self.current_window: Optional[WindowInfo] = None

    def find_window(
            self,
            title: Optional[str] = None,
            title_regex: Optional[str] = None,
            class_name: Optional[str] = None,
            process_name: Optional[str] = None,
            timeout: Optional[int] = None
    ) -> Optional[WindowInfo]:
        """
        Find a window by title, class, or process name.

        Args:
            title: Exact window title to match
            title_regex: Regex pattern for window title
            class_name: Window class name
            process_name: Process name (e.g., 'notepad.exe')
            timeout: Timeout in seconds (uses default if None)

        Returns:
            WindowInfo: Window information if found, None otherwise
        """
        timeout = timeout or self.default_timeout
        start_time = time.time()

        self.logger.info(f"Searching for window (timeout={timeout}s)...")

        while time.time() - start_time < timeout:
            windows = self.get_all_windows()

            for window in windows:
                # Check title match
                if title and window.title == title:
                    self.current_window = window
                    self.logger.info(f"Found window: {window}")
                    return window

                # Check title regex match
                if title_regex and re.search(title_regex, window.title):
                    self.current_window = window
                    self.logger.info(f"Found window by regex: {window}")
                    return window

                # Check class name match
                if class_name and window.class_name == class_name:
                    self.current_window = window
                    self.logger.info(f"Found window by class: {window}")
                    return window

                # Check process name match
                if process_name:
                    try:
                        process_path = self._get_process_path(window.process_id)
                        if process_name.lower() in process_path.lower():
                            self.current_window = window
                            self.logger.info(f"Found window by process: {window}")
                            return window
                    except:
                        pass

            # Wait before retry
            time.sleep(0.5)

        self.logger.warning("Window not found within timeout")
        return None

    def get_all_windows(self, visible_only: bool = True) -> List[WindowInfo]:
        """
        Get all windows currently open.

        Args:
            visible_only: Only return visible windows

        Returns:
            list: List of WindowInfo objects
        """
        windows = []

        def callback(hwnd, _):
            if visible_only and not win32gui.IsWindowVisible(hwnd):
                return

            title = win32gui.GetWindowText(hwnd)
            if not title:
                return

            try:
                class_name = win32gui.GetClassName(hwnd)
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)

                windows.append(WindowInfo(hwnd, title, class_name, process_id))
            except:
                pass

        win32gui.EnumWindows(callback, None)
        return windows

    def activate_window(self, window: Optional[WindowInfo] = None) -> bool:
        """
        Activate (bring to foreground) a window.

        Args:
            window: Window to activate (uses current window if None)

        Returns:
            bool: True if successful
        """
        window = window or self.current_window

        if not window:
            self.logger.error("No window specified or current window set")
            return False

        try:
            # Restore if minimized
            if win32gui.IsIconic(window.hwnd):
                win32gui.ShowWindow(window.hwnd, win32con.SW_RESTORE)

            # Bring to foreground
            win32gui.SetForegroundWindow(window.hwnd)

            # Wait a bit for window to activate
            time.sleep(0.2)

            self.logger.info(f"Activated window: {window.title}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to activate window: {e}")
            return False

    def get_window_rect(self, window: Optional[WindowInfo] = None) -> Optional[Tuple[int, int, int, int]]:
        """
        Get window rectangle (left, top, right, bottom).

        Args:
            window: Window to query (uses current window if None)

        Returns:
            tuple: (left, top, right, bottom) or None
        """
        window = window or self.current_window

        if not window:
            return None

        try:
            return win32gui.GetWindowRect(window.hwnd)
        except Exception as e:
            self.logger.error(f"Failed to get window rect: {e}")
            return None

    def set_window_position(
            self,
            x: int,
            y: int,
            width: int,
            height: int,
            window: Optional[WindowInfo] = None
    ) -> bool:
        """
        Set window position and size.

        Args:
            x: X position
            y: Y position
            width: Window width
            height: Window height
            window: Window to position (uses current window if None)

        Returns:
            bool: True if successful
        """
        window = window or self.current_window

        if not window:
            return False

        try:
            win32gui.SetWindowPos(
                window.hwnd,
                win32con.HWND_TOP,
                x, y, width, height,
                win32con.SWP_SHOWWINDOW
            )
            self.logger.info(f"Positioned window at ({x}, {y}) with size ({width}, {height})")
            return True

        except Exception as e:
            self.logger.error(f"Failed to set window position: {e}")
            return False

    def maximize_window(self, window: Optional[WindowInfo] = None) -> bool:
        """
        Maximize a window.

        Args:
            window: Window to maximize (uses current window if None)

        Returns:
            bool: True if successful
        """
        window = window or self.current_window

        if not window:
            return False

        try:
            win32gui.ShowWindow(window.hwnd, win32con.SW_MAXIMIZE)
            self.logger.info("Window maximized")
            return True
        except Exception as e:
            self.logger.error(f"Failed to maximize window: {e}")
            return False

    def minimize_window(self, window: Optional[WindowInfo] = None) -> bool:
        """
        Minimize a window.

        Args:
            window: Window to minimize (uses current window if None)

        Returns:
            bool: True if successful
        """
        window = window or self.current_window

        if not window:
            return False

        try:
            win32gui.ShowWindow(window.hwnd, win32con.SW_MINIMIZE)
            self.logger.info("Window minimized")
            return True
        except Exception as e:
            self.logger.error(f"Failed to minimize window: {e}")
            return False

    def close_window(self, window: Optional[WindowInfo] = None) -> bool:
        """
        Close a window.

        Args:
            window: Window to close (uses current window if None)

        Returns:
            bool: True if successful
        """
        window = window or self.current_window

        if not window:
            return False

        try:
            win32gui.PostMessage(window.hwnd, win32con.WM_CLOSE, 0, 0)
            self.logger.info("Sent close message to window")
            return True
        except Exception as e:
            self.logger.error(f"Failed to close window: {e}")
            return False

    def wait_for_window(
            self,
            title: Optional[str] = None,
            title_regex: Optional[str] = None,
            timeout: Optional[int] = None
    ) -> Optional[WindowInfo]:
        """
        Wait for a window to appear.

        Args:
            title: Window title to wait for
            title_regex: Regex pattern for window title
            timeout: Timeout in seconds

        Returns:
            WindowInfo: Window information if found
        """
        return self.find_window(title=title, title_regex=title_regex, timeout=timeout)

    def is_window_visible(self, window: Optional[WindowInfo] = None) -> bool:
        """
        Check if a window is visible.

        Args:
            window: Window to check (uses current window if None)

        Returns:
            bool: True if visible
        """
        window = window or self.current_window

        if not window:
            return False

        try:
            return bool(win32gui.IsWindowVisible(window.hwnd))
        except:
            return False

    def get_window_title(self, window: Optional[WindowInfo] = None) -> Optional[str]:
        """
        Get current window title (may have changed since window was found).

        Args:
            window: Window to query (uses current window if None)

        Returns:
            str: Window title or None
        """
        window = window or self.current_window

        if not window:
            return None

        try:
            return win32gui.GetWindowText(window.hwnd)
        except:
            return None

    @staticmethod
    def _get_process_path(process_id: int) -> str:
        """
        Get the executable path for a process ID.

        Args:
            process_id: Process ID

        Returns:
            str: Process executable path
        """
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, process_id)
            exe_path = win32process.GetModuleFileNameEx(handle, 0)
            win32api.CloseHandle(handle)
            return exe_path
        except:
            return ""

    def get_active_window(self) -> Optional[WindowInfo]:
        """
        Get currently active (foreground) window.

        Returns:
            WindowInfo: Active window information or None
        """
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)

            return WindowInfo(hwnd, title, class_name, process_id)
        except Exception as e:
            self.logger.error(f"Failed to get active window: {e}")
            return None
