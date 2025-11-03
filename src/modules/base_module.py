"""
Base module template for all automation modules.
Provides common functionality, logging, configuration, and error handling.
"""

import logging
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional
from datetime import datetime
from pathlib import Path
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.config import get_config_manager
from core.logging_config import get_logging_manager
from core.error_handler import ErrorCategory, ErrorSeverity, get_error_handler


class ModuleStatus:
    """Module execution status constants."""
    NOT_CONFIGURED = "not_configured"
    READY = "ready"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class BaseModule(ABC):
    """
    Base class for all automation modules.
    Provides common functionality for configuration, logging, and error handling.
    """

    def __init__(
            self,
            name: str,
            description: str,
            version: str = "1.0.0",
            category: str = "general"
    ):
        """
        Initialize the base module.

        Args:
            name: Module name
            description: Module description
            version: Module version
            category: Module category
        """
        self.name = name
        self.description = description
        self.version = version
        self.category = category

        # Get managers
        self.config_manager = get_config_manager()
        self.logging_manager = get_logging_manager()
        self.error_handler = get_error_handler()

        # Get module-specific logger
        self.logger = self.logging_manager.get_module_logger(name)

        # Module state
        self.status = ModuleStatus.NOT_CONFIGURED
        self.start_time: Optional[datetime] = None
        self.end_time: Optional[datetime] = None
        self.error: Optional[Exception] = None
        self.result: Optional[Dict[str, Any]] = None

        # Module configuration
        self._config: Dict[str, Any] = {}

        self.logger.info(f"Module initialized: {name} v{version}")

    @abstractmethod
    def configure(self, **kwargs) -> bool:
        """
        Configure the module with parameters.

        Args:
            **kwargs: Configuration parameters

        Returns:
            bool: True if configuration successful

        Note:
            This method must be implemented by subclasses.
        """
        pass

    @abstractmethod
    def validate(self) -> bool:
        """
        Validate configuration before execution.

        Returns:
            bool: True if configuration is valid

        Note:
            This method must be implemented by subclasses.
        """
        pass

    @abstractmethod
    def execute(self) -> Dict[str, Any]:
        """
        Execute the automation.

        Returns:
            dict: Execution result with status and data

        Note:
            This method must be implemented by subclasses.
        """
        pass

    def run(self, **config_kwargs) -> Dict[str, Any]:
        """
        Main entry point to configure and run the module.

        Args:
            **config_kwargs: Configuration parameters

        Returns:
            dict: Execution result
        """
        try:
            # Configure
            self.log_info("Configuring module...")
            if not self.configure(**config_kwargs):
                return self.error_result("Configuration failed")

            # Validate
            self.log_info("Validating configuration...")
            if not self.validate():
                return self.error_result("Validation failed")

            # Update status
            self.status = ModuleStatus.READY
            self.log_info("Configuration complete, starting execution...")

            # Execute
            self.status = ModuleStatus.RUNNING
            self.start_time = datetime.now()

            result = self.execute()

            self.end_time = datetime.now()
            self.status = ModuleStatus.COMPLETED
            self.result = result

            duration = (self.end_time - self.start_time).total_seconds()
            self.log_info(f"Execution completed in {duration:.2f} seconds")

            return result

        except Exception as e:
            self.status = ModuleStatus.FAILED
            self.error = e
            self.end_time = datetime.now()

            error_info = self.handle_error(
                e,
                context=f"Executing {self.name}",
                category=ErrorCategory.AUTOMATION
            )

            return self.error_result(error_info['message'])

    # Logging methods
    def log_info(self, message: str):
        """Log an info message."""
        self.logger.info(f"[{self.name}] {message}")

    def log_warning(self, message: str):
        """Log a warning message."""
        self.logger.warning(f"[{self.name}] {message}")

    def log_error(self, message: str):
        """Log an error message."""
        self.logger.error(f"[{self.name}] {message}")

    def log_debug(self, message: str):
        """Log a debug message."""
        self.logger.debug(f"[{self.name}] {message}")

    # Configuration methods
    def get_config(self, key: str, default: Any = None) -> Any:
        """
        Get a module configuration value.

        Args:
            key: Configuration key
            default: Default value if not found

        Returns:
            Configuration value or default
        """
        return self._config.get(key, default)

    def set_config(self, key: str, value: Any):
        """
        Set a module configuration value.

        Args:
            key: Configuration key
            value: Configuration value
        """
        self._config[key] = value

    def get_app_config(self, key: str, default: Any = None) -> Any:
        """
        Get an application configuration value.

        Args:
            key: Configuration key (dot notation)
            default: Default value

        Returns:
            Configuration value or default
        """
        return self.config_manager.get(key, default)

    def save_module_config(self):
        """Save module configuration to persistent storage."""
        self.config_manager.set_module_config(self.name, self._config)

    def load_module_config(self):
        """Load module configuration from persistent storage."""
        self._config = self.config_manager.get_module_config(self.name, {})

    # Error handling methods
    def handle_error(
            self,
            exception: Exception,
            context: Optional[str] = None,
            category: ErrorCategory = ErrorCategory.AUTOMATION,
            severity: ErrorSeverity = ErrorSeverity.MEDIUM
    ) -> Dict[str, Any]:
        """
        Handle an error with logging and tracking.

        Args:
            exception: The exception
            context: Context description
            category: Error category
            severity: Error severity

        Returns:
            dict: Error information
        """
        return self.error_handler.handle_error(
            exception,
            context=context or f"{self.name} module",
            category=category,
            severity=severity
        )

    # Result helpers
    @staticmethod
    def success_result(data: Any = None, message: str = "Success") -> Dict[str, Any]:
        """
        Create a success result.

        Args:
            data: Result data
            message: Success message

        Returns:
            dict: Success result
        """
        return {
            'status': 'success',
            'message': message,
            'data': data,
            'timestamp': datetime.now().isoformat()
        }

    @staticmethod
    def error_result(message: str, error_data: Any = None) -> Dict[str, Any]:
        """
        Create an error result.

        Args:
            message: Error message
            error_data: Additional error data

        Returns:
            dict: Error result
        """
        return {
            'status': 'error',
            'message': message,
            'error': error_data,
            'timestamp': datetime.now().isoformat()
        }

    # Utility methods
    def get_output_path(self, filename: str) -> Path:
        """
        Get the full path for an output file.

        Args:
            filename: Output filename

        Returns:
            Path: Full output path
        """
        output_dir = self.get_app_config('paths.output')
        output_path = Path(output_dir) / self.name / filename
        output_path.parent.mkdir(parents=True, exist_ok=True)
        return output_path

    def get_temp_path(self, filename: str) -> Path:
        """
        Get the full path for a temporary file.

        Args:
            filename: Temp filename

        Returns:
            Path: Full temp path
        """
        temp_dir = self.get_app_config('paths.temp')
        temp_path = Path(temp_dir) / self.name / filename
        temp_path.parent.mkdir(parents=True, exist_ok=True)
        return temp_path

    def get_duration(self) -> Optional[float]:
        """
        Get execution duration in seconds.

        Returns:
            float: Duration in seconds or None if not completed
        """
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return None

    def get_status_info(self) -> Dict[str, Any]:
        """
        Get module status information.

        Returns:
            dict: Status information
        """
        return {
            'name': self.name,
            'description': self.description,
            'version': self.version,
            'category': self.category,
            'status': self.status,
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
            'duration': self.get_duration(),
            'error': str(self.error) if self.error else None
        }

    def reset(self):
        """Reset module state for re-execution."""
        self.status = ModuleStatus.NOT_CONFIGURED
        self.start_time = None
        self.end_time = None
        self.error = None
        self.result = None
        self.log_info("Module reset")

    def __str__(self) -> str:
        return f"{self.name} v{self.version} [{self.status}]"

    def __repr__(self) -> str:
        return f"BaseModule(name='{self.name}', version='{self.version}', status='{self.status}')"
