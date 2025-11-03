"""
Automation Hub - Main Entry Point
Launches the PyQt5 GUI application.
"""

import sys
import logging
from pathlib import Path

# Ensure the src directory is in the path
sys.path.insert(0, str(Path(__file__).parent))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from core.config import get_config_manager
from core.logging_config import get_logging_manager
from core.error_handler import get_error_handler
from hub.main_window import MainWindow


def setup_application():
    """Initialize application services."""
    # Initialize configuration
    config_manager = get_config_manager()

    # Initialize logging
    log_level = config_manager.get('app.log_level', 'INFO')
    logging_manager = get_logging_manager(log_level=log_level)

    # Initialize error handler
    error_handler = get_error_handler()

    logger = logging.getLogger(__name__)
    logger.info("=" * 80)
    logger.info("Automation Hub Starting")
    logger.info("=" * 80)

    return config_manager, logging_manager, error_handler


def main():
    """Main application entry point."""
    try:
        # Set up high DPI scaling
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

        # Create QApplication
        app = QApplication(sys.argv)
        app.setApplicationName("Automation Hub")
        app.setOrganizationName("AutomationHub")

        # Setup application services
        setup_application()

        # Create and show main window
        main_window = MainWindow()
        main_window.show()

        # Start event loop
        sys.exit(app.exec_())

    except Exception as e:
        logging.critical(f"Fatal error starting application: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
