"""
Input Controller for Desktop RPA.
Handles mouse and keyboard input automation using pyautogui.
"""

import logging
import time
from typing import Optional, Tuple, List
from pathlib import Path

try:
    import pyautogui
    # Set pyautogui safety settings
    pyautogui.PAUSE = 0.1  # Default pause between actions
    pyautogui.FAILSAFE = True  # Move mouse to corner to abort
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
    logging.warning("pyautogui not available - input automation features will be limited")


class InputController:
    """
    Controls mouse and keyboard input for RPA automation.
    Provides safe, reliable input simulation with configurable delays and retries.
    """

    def __init__(
            self,
            default_delay: float = 0.5,
            typing_interval: float = 0.05,
            mouse_duration: float = 0.2
    ):
        """
        Initialize the input controller.

        Args:
            default_delay: Default delay between actions in seconds
            typing_interval: Delay between keystrokes in seconds
            mouse_duration: Duration of mouse movements in seconds
        """
        if not PYAUTOGUI_AVAILABLE:
            raise RuntimeError("pyautogui is required for InputController")

        self.logger = logging.getLogger(__name__)
        self.default_delay = default_delay
        self.typing_interval = typing_interval
        self.mouse_duration = mouse_duration

        # Get screen size
        self.screen_width, self.screen_height = pyautogui.size()
        self.logger.info(f"Screen size: {self.screen_width}x{self.screen_height}")

    def click(
            self,
            x: Optional[int] = None,
            y: Optional[int] = None,
            button: str = 'left',
            clicks: int = 1,
            delay: Optional[float] = None
    ) -> bool:
        """
        Click at specified position or current mouse position.

        Args:
            x: X coordinate (None for current position)
            y: Y coordinate (None for current position)
            button: Mouse button ('left', 'right', 'middle')
            clicks: Number of clicks (1 for single, 2 for double)
            delay: Delay after click (uses default if None)

        Returns:
            bool: True if successful
        """
        try:
            if x is not None and y is not None:
                self.logger.debug(f"Clicking at ({x}, {y}) with {button} button, {clicks} times")
                pyautogui.click(x=x, y=y, clicks=clicks, button=button, duration=self.mouse_duration)
            else:
                self.logger.debug(f"Clicking at current position with {button} button, {clicks} times")
                pyautogui.click(clicks=clicks, button=button)

            # Delay after click
            time.sleep(delay or self.default_delay)
            return True

        except Exception as e:
            self.logger.error(f"Click failed: {e}")
            return False

    def double_click(
            self,
            x: Optional[int] = None,
            y: Optional[int] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Double-click at specified position.

        Args:
            x: X coordinate
            y: Y coordinate
            delay: Delay after click

        Returns:
            bool: True if successful
        """
        return self.click(x=x, y=y, clicks=2, delay=delay)

    def right_click(
            self,
            x: Optional[int] = None,
            y: Optional[int] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Right-click at specified position.

        Args:
            x: X coordinate
            y: Y coordinate
            delay: Delay after click

        Returns:
            bool: True if successful
        """
        return self.click(x=x, y=y, button='right', delay=delay)

    def move_to(
            self,
            x: int,
            y: int,
            duration: Optional[float] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Move mouse to specified position.

        Args:
            x: X coordinate
            y: Y coordinate
            duration: Movement duration in seconds
            delay: Delay after movement

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Moving mouse to ({x}, {y})")
            pyautogui.moveTo(x, y, duration=duration or self.mouse_duration)
            time.sleep(delay or (self.default_delay / 2))
            return True

        except Exception as e:
            self.logger.error(f"Move failed: {e}")
            return False

    def move_relative(
            self,
            x_offset: int,
            y_offset: int,
            duration: Optional[float] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Move mouse relative to current position.

        Args:
            x_offset: X offset
            y_offset: Y offset
            duration: Movement duration in seconds
            delay: Delay after movement

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Moving mouse by ({x_offset}, {y_offset})")
            pyautogui.move(x_offset, y_offset, duration=duration or self.mouse_duration)
            time.sleep(delay or (self.default_delay / 2))
            return True

        except Exception as e:
            self.logger.error(f"Relative move failed: {e}")
            return False

    def drag_to(
            self,
            x: int,
            y: int,
            duration: Optional[float] = None,
            button: str = 'left',
            delay: Optional[float] = None
    ) -> bool:
        """
        Drag mouse to specified position.

        Args:
            x: X coordinate
            y: Y coordinate
            duration: Drag duration in seconds
            button: Mouse button to hold
            delay: Delay after drag

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Dragging to ({x}, {y})")
            pyautogui.dragTo(x, y, duration=duration or (self.mouse_duration * 2), button=button)
            time.sleep(delay or self.default_delay)
            return True

        except Exception as e:
            self.logger.error(f"Drag failed: {e}")
            return False

    def scroll(
            self,
            clicks: int,
            x: Optional[int] = None,
            y: Optional[int] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Scroll mouse wheel.

        Args:
            clicks: Number of scroll clicks (positive=up, negative=down)
            x: X coordinate to scroll at (None for current position)
            y: Y coordinate to scroll at (None for current position)
            delay: Delay after scroll

        Returns:
            bool: True if successful
        """
        try:
            if x is not None and y is not None:
                pyautogui.scroll(clicks, x=x, y=y)
            else:
                pyautogui.scroll(clicks)

            self.logger.debug(f"Scrolled {clicks} clicks")
            time.sleep(delay or (self.default_delay / 2))
            return True

        except Exception as e:
            self.logger.error(f"Scroll failed: {e}")
            return False

    def type_text(
            self,
            text: str,
            interval: Optional[float] = None,
            delay: Optional[float] = None
    ) -> bool:
        """
        Type text character by character.

        Args:
            text: Text to type
            interval: Interval between keystrokes
            delay: Delay after typing

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Typing text: {text[:50]}...")
            pyautogui.typewrite(text, interval=interval or self.typing_interval)
            time.sleep(delay or self.default_delay)
            return True

        except Exception as e:
            self.logger.error(f"Type text failed: {e}")
            return False

    def press_key(
            self,
            key: str,
            presses: int = 1,
            interval: float = 0.0,
            delay: Optional[float] = None
    ) -> bool:
        """
        Press a key one or more times.

        Args:
            key: Key name (e.g., 'enter', 'tab', 'esc', 'a', 'ctrl')
            presses: Number of times to press
            interval: Interval between presses
            delay: Delay after pressing

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Pressing key: {key} ({presses} times)")
            pyautogui.press(key, presses=presses, interval=interval)
            time.sleep(delay or (self.default_delay / 2))
            return True

        except Exception as e:
            self.logger.error(f"Press key failed: {e}")
            return False

    def hotkey(
            self,
            *keys: str,
            delay: Optional[float] = None
    ) -> bool:
        """
        Press a hotkey combination (e.g., Ctrl+C, Alt+Tab).

        Args:
            *keys: Keys to press together (e.g., 'ctrl', 'c')
            delay: Delay after hotkey

        Returns:
            bool: True if successful
        """
        try:
            self.logger.debug(f"Pressing hotkey: {'+'.join(keys)}")
            pyautogui.hotkey(*keys)
            time.sleep(delay or self.default_delay)
            return True

        except Exception as e:
            self.logger.error(f"Hotkey failed: {e}")
            return False

    def get_mouse_position(self) -> Tuple[int, int]:
        """
        Get current mouse position.

        Returns:
            tuple: (x, y) coordinates
        """
        return pyautogui.position()

    def screenshot(
            self,
            filepath: Optional[str] = None,
            region: Optional[Tuple[int, int, int, int]] = None
    ) -> Optional[str]:
        """
        Take a screenshot.

        Args:
            filepath: Path to save screenshot (auto-generated if None)
            region: Region to capture (left, top, width, height)

        Returns:
            str: Path to saved screenshot or None if failed
        """
        try:
            if filepath is None:
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filepath = f"screenshot_{timestamp}.png"

            if region:
                screenshot = pyautogui.screenshot(region=region)
            else:
                screenshot = pyautogui.screenshot()

            screenshot.save(filepath)
            self.logger.info(f"Screenshot saved to: {filepath}")
            return filepath

        except Exception as e:
            self.logger.error(f"Screenshot failed: {e}")
            return None

    def locate_on_screen(
            self,
            image_path: str,
            confidence: float = 0.9,
            grayscale: bool = True
    ) -> Optional[Tuple[int, int, int, int]]:
        """
        Locate an image on screen.

        Args:
            image_path: Path to image file
            confidence: Matching confidence (0.0 to 1.0)
            grayscale: Convert to grayscale for faster matching

        Returns:
            tuple: (left, top, width, height) of found image or None
        """
        try:
            location = pyautogui.locateOnScreen(
                image_path,
                confidence=confidence,
                grayscale=grayscale
            )

            if location:
                self.logger.info(f"Image found at: {location}")
                return location
            else:
                self.logger.debug("Image not found on screen")
                return None

        except Exception as e:
            self.logger.error(f"Locate on screen failed: {e}")
            return None

    def click_on_image(
            self,
            image_path: str,
            confidence: float = 0.9,
            button: str = 'left',
            timeout: int = 10,
            delay: Optional[float] = None
    ) -> bool:
        """
        Find and click on an image.

        Args:
            image_path: Path to image file
            confidence: Matching confidence
            button: Mouse button
            timeout: Timeout in seconds
            delay: Delay after click

        Returns:
            bool: True if found and clicked
        """
        start_time = time.time()

        while time.time() - start_time < timeout:
            location = self.locate_on_screen(image_path, confidence=confidence)

            if location:
                # Click center of found image
                center_x = location[0] + location[2] // 2
                center_y = location[1] + location[3] // 2
                return self.click(center_x, center_y, button=button, delay=delay)

            time.sleep(0.5)

        self.logger.warning(f"Image not found within timeout: {image_path}")
        return False

    def wait(self, seconds: float):
        """
        Wait for specified duration.

        Args:
            seconds: Duration to wait in seconds
        """
        self.logger.debug(f"Waiting {seconds} seconds")
        time.sleep(seconds)

    def is_failsafe_triggered(self) -> bool:
        """
        Check if failsafe is triggered (mouse in corner).

        Returns:
            bool: True if failsafe position
        """
        x, y = self.get_mouse_position()
        return x == 0 and y == 0

    def set_typing_speed(self, interval: float):
        """
        Set typing speed.

        Args:
            interval: Interval between keystrokes in seconds
        """
        self.typing_interval = interval
        self.logger.info(f"Typing interval set to {interval} seconds")

    def set_mouse_speed(self, duration: float):
        """
        Set mouse movement speed.

        Args:
            duration: Duration for mouse movements in seconds
        """
        self.mouse_duration = duration
        self.logger.info(f"Mouse duration set to {duration} seconds")

    def set_default_delay(self, delay: float):
        """
        Set default delay between actions.

        Args:
            delay: Delay in seconds
        """
        self.default_delay = delay
        pyautogui.PAUSE = delay
        self.logger.info(f"Default delay set to {delay} seconds")
