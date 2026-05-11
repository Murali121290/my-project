"""
Logging configuration for Book Styler application.
Provides consistent logging across all modules.
"""

import logging
import sys
from pathlib import Path
from typing import Optional

# Lazy import to avoid circular imports



class ColoredFormatter(logging.Formatter):
    """Custom formatter with color support for console output"""
    
    COLORS = {
        'DEBUG': '\033[36m',      # Cyan
        'INFO': '\033[32m',       # Green
        'WARNING': '\033[33m',    # Yellow
        'ERROR': '\033[31m',      # Red
        'CRITICAL': '\033[35m',   # Magenta
    }
    RESET = '\033[0m'
    
    def format(self, record):
        if hasattr(sys.stderr, 'isatty') and sys.stderr.isatty():
            levelname = record.levelname
            record.levelname = f"{self.COLORS.get(levelname, '')}{levelname}{self.RESET}"
        return super().format(record)


def setup_logging(name: Optional[str] = None, level: Optional[str] = None) -> logging.Logger:
    """
    Setup and return a logger instance.
    
    Args:
        name: Logger name (defaults to root logger)
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
               Defaults to config LOG_LEVEL
    
    Returns:
        Configured logger instance
    """
    if level is None:
        level = "INFO"
    
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, level, logging.INFO))
    
    # Avoid duplicate handlers
    if logger.handlers:
        return logger
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, level))
    console_formatter = ColoredFormatter(
        fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # File handler
    # File handler (Optional - could use app config, but for now just console)
    # log_file = Path("app.log")
    # if log_file.parent.exists():
    #     file_handler = logging.FileHandler(log_file)
    #     file_handler.setLevel(getattr(logging, level, logging.INFO))
    #     file_formatter = logging.Formatter(
    #         fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    #         datefmt='%Y-%m-%d %H:%M:%S'
    #     )
    #     file_handler.setFormatter(file_formatter)
    #     logger.addHandler(file_handler)
    
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Get or create a logger with the given name"""
    if name is None:
        return logging.getLogger()
    
    logger = logging.getLogger(name)
    if not logger.handlers:
        setup_logging(name)
    
    return logger
