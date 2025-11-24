import os
import sys
from pathlib import Path
from loguru import logger as loguru_logger

from config import settings


def setup_logger():
    log_dir = settings.LOG_FILE.parent
    os.makedirs(log_dir, exist_ok=True)
    
    loguru_logger.remove()
    
    log_format = (
        "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
        "<level>{level: <8}</level> | "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - "
        "<level>{message}</level>"
    )
    
    loguru_logger.add(
        sys.stderr,
        format=log_format,
        level=settings.LOG_LEVEL,
        colorize=True
    )
    
    loguru_logger.add(
        str(settings.LOG_FILE),
        format=log_format,
        level=settings.LOG_LEVEL,
        rotation="10 MB",
        retention="30 days",
        compression="zip",
        encoding="utf-8"
    )
    
    return loguru_logger


logger = setup_logger()