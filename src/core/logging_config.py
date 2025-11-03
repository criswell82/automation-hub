"""
Centralized logging framework for Automation Hub.
Provides structured logging with file rotation, multiple handlers, and filtering.
"""

import logging
import logging.handlers
import os
from pathlib import Path
from typing import Optional
from datetime import datetime


class LoggingManager:
    """
    Manages application-wide logging with multiple handlers and rotation.
    """

    def __init__(self, log_dir: Optional[str] = None, log_level: str = 'INFO'):
        """
        Initialize the logging manager.

        Args:
            log_dir: Directory for log files. Defaults to AppData/AutomationHub/logs
            log_level: Default logging level ('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')
        """
        if log_dir is None:
            appdata = os.getenv('APPDATA', os.path.expanduser('~'))
            log_dir = os.path.join(appdata, 'AutomationHub', 'logs')

        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        self.log_level = getattr(logging, log_level.upper(), logging.INFO)

        # Log files
        self.main_log_file = self.log_dir / 'automation_hub.log'
        self.error_log_file = self.log_dir / 'errors.log'
        self.module_log_dir = self.log_dir / 'modules'
        self.module_log_dir.mkdir(exist_ok=True)

        # Configure root logger
        self._setup_root_logger()

    def _setup_root_logger(self):
        """Configure the root logger with handlers."""
        root_logger = logging.getLogger()
        root_logger.setLevel(self.log_level)

        # Remove any existing handlers
        root_logger.handlers.clear()

        # Create formatters
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        simple_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(self.log_level)
        console_handler.setFormatter(simple_formatter)
        root_logger.addHandler(console_handler)

        # Main log file handler with rotation (10 MB per file, keep 5 backups)
        file_handler = logging.handlers.RotatingFileHandler(
            self.main_log_file,
            maxBytes=10 * 1024 * 1024,  # 10 MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(self.log_level)
        file_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(file_handler)

        # Error log file handler (only for ERROR and above)
        error_handler = logging.handlers.RotatingFileHandler(
            self.error_log_file,
            maxBytes=5 * 1024 * 1024,  # 5 MB
            backupCount=3,
            encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(error_handler)

        # Log startup message
        logging.info("=" * 80)
        logging.info(f"Automation Hub Logging Started - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logging.info(f"Log Level: {logging.getLevelName(self.log_level)}")
        logging.info(f"Log Directory: {self.log_dir}")
        logging.info("=" * 80)

    def get_logger(self, name: str) -> logging.Logger:
        """
        Get a logger instance for a specific module or component.

        Args:
            name: Name of the logger (typically module name)

        Returns:
            Logger instance
        """
        logger = logging.getLogger(name)
        logger.setLevel(self.log_level)
        return logger

    def get_module_logger(self, module_name: str) -> logging.Logger:
        """
        Get a logger for a specific automation module with its own log file.

        Args:
            module_name: Name of the automation module

        Returns:
            Logger instance with dedicated file handler
        """
        logger_name = f"module.{module_name}"
        logger = logging.getLogger(logger_name)
        logger.setLevel(self.log_level)

        # Check if this logger already has a file handler
        has_file_handler = any(
            isinstance(h, logging.handlers.RotatingFileHandler)
            for h in logger.handlers
        )

        if not has_file_handler:
            # Add a dedicated file handler for this module
            module_log_file = self.module_log_dir / f"{module_name}.log"

            module_handler = logging.handlers.RotatingFileHandler(
                module_log_file,
                maxBytes=5 * 1024 * 1024,  # 5 MB
                backupCount=3,
                encoding='utf-8'
            )
            module_handler.setLevel(self.log_level)

            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            module_handler.setFormatter(formatter)

            logger.addHandler(module_handler)

            # Don't propagate to root logger to avoid duplicate logs
            logger.propagate = False

            # But add a console handler for visibility
            console_handler = logging.StreamHandler()
            console_handler.setLevel(self.log_level)
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

        return logger

    def set_log_level(self, level: str):
        """
        Change the logging level for all loggers.

        Args:
            level: New logging level ('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')
        """
        self.log_level = getattr(logging, level.upper(), logging.INFO)

        root_logger = logging.getLogger()
        root_logger.setLevel(self.log_level)

        for handler in root_logger.handlers:
            handler.setLevel(self.log_level)

        logging.info(f"Log level changed to: {level.upper()}")

    def cleanup_old_logs(self, days: int = 30):
        """
        Delete log files older than specified days.

        Args:
            days: Delete logs older than this many days
        """
        import time
        from datetime import timedelta

        cutoff_time = time.time() - (days * 86400)  # days to seconds

        deleted_count = 0

        for log_file in self.log_dir.rglob('*.log*'):
            if log_file.is_file():
                if log_file.stat().st_mtime < cutoff_time:
                    try:
                        log_file.unlink()
                        deleted_count += 1
                    except Exception as e:
                        logging.warning(f"Could not delete old log file {log_file}: {e}")

        logging.info(f"Cleaned up {deleted_count} old log files (older than {days} days)")

    def get_recent_logs(self, lines: int = 100, log_file: Optional[str] = None) -> list:
        """
        Get the most recent log entries.

        Args:
            lines: Number of recent lines to retrieve
            log_file: Specific log file to read (defaults to main log)

        Returns:
            List of log lines
        """
        if log_file is None:
            log_file = self.main_log_file

        log_path = Path(log_file) if '/' in log_file or '\\' in log_file else self.log_dir / log_file

        try:
            if not log_path.exists():
                return []

            with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                # Read last N lines efficiently
                return self._tail_file(f, lines)

        except Exception as e:
            logging.error(f"Could not read log file {log_path}: {e}")
            return []

    @staticmethod
    def _tail_file(file_handle, n: int) -> list:
        """Read last N lines from a file efficiently."""
        # Simple implementation - for large files, could use more efficient algorithm
        lines = file_handle.readlines()
        return lines[-n:] if len(lines) > n else lines

    def archive_logs(self, archive_name: Optional[str] = None):
        """
        Archive current logs to a zip file.

        Args:
            archive_name: Name of the archive file (defaults to timestamp-based name)
        """
        import zipfile

        if archive_name is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            archive_name = f"logs_archive_{timestamp}.zip"

        archive_path = self.log_dir / archive_name

        try:
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for log_file in self.log_dir.rglob('*.log*'):
                    if log_file.is_file() and log_file != archive_path:
                        arcname = log_file.relative_to(self.log_dir)
                        zipf.write(log_file, arcname)

            logging.info(f"Logs archived to: {archive_path}")
            return str(archive_path)

        except Exception as e:
            logging.error(f"Failed to archive logs: {e}")
            return None


# Singleton instance
_logging_manager: Optional[LoggingManager] = None


def get_logging_manager(log_dir: Optional[str] = None, log_level: str = 'INFO') -> LoggingManager:
    """Get the global LoggingManager instance."""
    global _logging_manager
    if _logging_manager is None:
        _logging_manager = LoggingManager(log_dir, log_level)
    return _logging_manager


def get_logger(name: str) -> logging.Logger:
    """
    Convenience function to get a logger.

    Args:
        name: Logger name (typically __name__)

    Returns:
        Logger instance
    """
    manager = get_logging_manager()
    return manager.get_logger(name)
