"""
Logging Configuration Module
"""

import logging
import logging.handlers
import os
from datetime import datetime

def setup_logging(log_level='INFO', log_dir='logs', max_files=5, max_size_mb=10):
    """
    Setup logging configuration
    
    Args:
        log_level (str): Logging level
        log_dir (str): Directory for log files
        max_files (int): Maximum number of log files to keep
        max_size_mb (int): Maximum size of each log file in MB
    """
    
    # Create logs directory if it doesn't exist
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    # Configure log filename with timestamp
    log_filename = os.path.join(log_dir, f'backorder_report_{datetime.now().strftime("%Y%m%d")}.log')
    
    # Create formatters
    detailed_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    simple_formatter = logging.Formatter(
        '%(levelname)s: %(message)s'
    )
    
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    
    # Remove any existing handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # File handler with rotation
    file_handler = logging.handlers.RotatingFileHandler(
        log_filename,
        maxBytes=max_size_mb * 1024 * 1024,  # Convert MB to bytes
        backupCount=max_files - 1
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(detailed_formatter)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(simple_formatter)
    
    # Add handlers to root logger
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
    
    # Create application logger
    app_logger = logging.getLogger('backorder_report')
    app_logger.info("Logging system initialized")
    
    return app_logger

class GUILogHandler(logging.Handler):
    """Custom log handler to display messages in GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        """Emit a log record to the GUI text widget"""
        try:
            msg = self.format(record)
            # Schedule the GUI update in the main thread
            self.text_widget.after(0, lambda: self._update_text(msg))
        except Exception:
            self.handleError(record)
            
    def _update_text(self, message):
        """Update the text widget with the log message"""
        try:
            self.text_widget.insert('end', message + '\n')
            self.text_widget.see('end')
        except Exception:
            pass  # Ignore errors when GUI is closing
