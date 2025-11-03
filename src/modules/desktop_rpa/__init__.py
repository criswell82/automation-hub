"""
Desktop RPA (Robotic Process Automation) module.
Provides UI automation capabilities for interacting with Windows applications.
"""

from .window_manager import WindowManager
from .input_controller import InputController

__all__ = ['WindowManager', 'InputController']
