"""
Centralized error handling for Automation Hub.
Provides structured error handling, recovery mechanisms, and user-friendly error messages.
"""

import logging
import traceback
import sys
from typing import Optional, Callable, Any, Dict
from enum import Enum
from datetime import datetime


class ErrorSeverity(Enum):
    """Error severity levels."""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


class ErrorCategory(Enum):
    """Error categories for better classification."""
    CONFIGURATION = "configuration"
    AUTHENTICATION = "authentication"
    NETWORK = "network"
    FILE_IO = "file_io"
    PERMISSION = "permission"
    VALIDATION = "validation"
    AUTOMATION = "automation"
    UNKNOWN = "unknown"


class AutomationError(Exception):
    """Base exception class for Automation Hub."""

    def __init__(
            self,
            message: str,
            category: ErrorCategory = ErrorCategory.UNKNOWN,
            severity: ErrorSeverity = ErrorSeverity.MEDIUM,
            original_exception: Optional[Exception] = None,
            context: Optional[Dict[str, Any]] = None
    ):
        """
        Initialize an AutomationError.

        Args:
            message: Error message
            category: Error category
            severity: Error severity
            original_exception: Original exception if this wraps another error
            context: Additional context information
        """
        super().__init__(message)
        self.message = message
        self.category = category
        self.severity = severity
        self.original_exception = original_exception
        self.context = context or {}
        self.timestamp = datetime.now()


class ErrorHandler:
    """
    Centralized error handling with logging, recovery, and user notifications.
    """

    def __init__(self):
        """Initialize the error handler."""
        self.logger = logging.getLogger(__name__)
        self.error_history = []
        self.max_history = 100
        self.error_callbacks = []

    def handle_error(
            self,
            exception: Exception,
            context: Optional[str] = None,
            category: ErrorCategory = ErrorCategory.UNKNOWN,
            severity: ErrorSeverity = ErrorSeverity.MEDIUM,
            user_message: Optional[str] = None,
            reraise: bool = False
    ) -> Dict[str, Any]:
        """
        Handle an exception with logging and optional recovery.

        Args:
            exception: The exception to handle
            context: Context description (e.g., "Loading configuration")
            category: Error category
            severity: Error severity
            user_message: User-friendly message (if None, generates from exception)
            reraise: Whether to re-raise the exception after handling

        Returns:
            dict: Error information including message, category, severity
        """
        # Generate user-friendly message if not provided
        if user_message is None:
            user_message = self._generate_user_message(exception, context)

        # Log the error
        log_message = f"{context}: {str(exception)}" if context else str(exception)

        if severity == ErrorSeverity.CRITICAL:
            self.logger.critical(log_message, exc_info=True)
        elif severity == ErrorSeverity.HIGH:
            self.logger.error(log_message, exc_info=True)
        elif severity == ErrorSeverity.MEDIUM:
            self.logger.warning(log_message, exc_info=True)
        else:  # LOW
            self.logger.info(log_message)

        # Create error info
        error_info = {
            'message': user_message,
            'technical_message': str(exception),
            'category': category.value,
            'severity': severity.value,
            'context': context,
            'timestamp': datetime.now().isoformat(),
            'traceback': traceback.format_exc() if sys.exc_info()[0] is not None else None
        }

        # Add to history
        self._add_to_history(error_info)

        # Notify callbacks
        self._notify_callbacks(error_info)

        # Re-raise if requested
        if reraise:
            raise

        return error_info

    def handle_automation_error(self, error: AutomationError) -> Dict[str, Any]:
        """
        Handle an AutomationError specifically.

        Args:
            error: AutomationError instance

        Returns:
            dict: Error information
        """
        return self.handle_error(
            error.original_exception or error,
            context=str(error.context) if error.context else None,
            category=error.category,
            severity=error.severity,
            user_message=error.message
        )

    def wrap_function(
            self,
            func: Callable,
            context: Optional[str] = None,
            category: ErrorCategory = ErrorCategory.UNKNOWN,
            default_return: Any = None
    ) -> Callable:
        """
        Wrap a function with automatic error handling.

        Args:
            func: Function to wrap
            context: Context description
            category: Error category
            default_return: Value to return on error

        Returns:
            Wrapped function
        """

        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                self.handle_error(e, context=context or func.__name__, category=category)
                return default_return

        return wrapper

    def safe_execute(
            self,
            func: Callable,
            *args,
            context: Optional[str] = None,
            category: ErrorCategory = ErrorCategory.UNKNOWN,
            default_return: Any = None,
            **kwargs
    ) -> Any:
        """
        Safely execute a function with error handling.

        Args:
            func: Function to execute
            *args: Positional arguments for func
            context: Context description
            category: Error category
            default_return: Value to return on error
            **kwargs: Keyword arguments for func

        Returns:
            Function result or default_return on error
        """
        try:
            return func(*args, **kwargs)
        except Exception as e:
            self.handle_error(e, context=context or func.__name__, category=category)
            return default_return

    def add_callback(self, callback: Callable[[Dict[str, Any]], None]):
        """
        Add a callback to be notified of errors.

        Args:
            callback: Function to call with error info dict
        """
        self.error_callbacks.append(callback)

    def get_error_history(self, limit: Optional[int] = None) -> list:
        """
        Get recent error history.

        Args:
            limit: Maximum number of errors to return

        Returns:
            List of error info dicts
        """
        if limit:
            return self.error_history[-limit:]
        return self.error_history.copy()

    def clear_error_history(self):
        """Clear the error history."""
        self.error_history.clear()
        self.logger.info("Error history cleared")

    def get_error_stats(self) -> Dict[str, Any]:
        """
        Get statistics about errors.

        Returns:
            dict: Error statistics
        """
        if not self.error_history:
            return {
                'total_errors': 0,
                'by_category': {},
                'by_severity': {}
            }

        by_category = {}
        by_severity = {}

        for error in self.error_history:
            category = error.get('category', 'unknown')
            severity = error.get('severity', 'medium')

            by_category[category] = by_category.get(category, 0) + 1
            by_severity[severity] = by_severity.get(severity, 0) + 1

        return {
            'total_errors': len(self.error_history),
            'by_category': by_category,
            'by_severity': by_severity
        }

    def _generate_user_message(self, exception: Exception, context: Optional[str]) -> str:
        """
        Generate a user-friendly error message.

        Args:
            exception: The exception
            context: Context description

        Returns:
            User-friendly message
        """
        error_type = type(exception).__name__
        error_msg = str(exception)

        # Common error patterns and user-friendly messages
        if isinstance(exception, FileNotFoundError):
            return f"File not found: {error_msg}. Please check the file path and try again."

        elif isinstance(exception, PermissionError):
            return f"Permission denied: {error_msg}. Please check file permissions or run with appropriate privileges."

        elif isinstance(exception, ConnectionError):
            return f"Connection failed: {error_msg}. Please check your network connection and try again."

        elif isinstance(exception, TimeoutError):
            return "The operation timed out. Please try again or increase the timeout value."

        elif isinstance(exception, ValueError):
            return f"Invalid value: {error_msg}. Please check your input and try again."

        elif isinstance(exception, KeyError):
            return f"Required configuration key missing: {error_msg}."

        # Default message
        if context:
            return f"An error occurred while {context}: {error_msg}"
        else:
            return f"An error occurred: {error_msg}"

    def _add_to_history(self, error_info: Dict[str, Any]):
        """Add error to history with size limit."""
        self.error_history.append(error_info)

        # Keep history size manageable
        if len(self.error_history) > self.max_history:
            self.error_history = self.error_history[-self.max_history:]

    def _notify_callbacks(self, error_info: Dict[str, Any]):
        """Notify registered callbacks of an error."""
        for callback in self.error_callbacks:
            try:
                callback(error_info)
            except Exception as e:
                self.logger.error(f"Error in error callback: {e}")


# Singleton instance
_error_handler: Optional[ErrorHandler] = None


def get_error_handler() -> ErrorHandler:
    """Get the global ErrorHandler instance."""
    global _error_handler
    if _error_handler is None:
        _error_handler = ErrorHandler()
    return _error_handler


# Convenience functions
def handle_error(exception: Exception, context: Optional[str] = None, **kwargs) -> Dict[str, Any]:
    """Convenience function to handle an error."""
    handler = get_error_handler()
    return handler.handle_error(exception, context=context, **kwargs)


def safe_execute(func: Callable, *args, **kwargs) -> Any:
    """Convenience function to safely execute a function."""
    handler = get_error_handler()
    return handler.safe_execute(func, *args, **kwargs)
