"""
Logging configuration using loguru.
One rotating file per day, plus console output during development.
Call setup_logging() once at startup from main.py.
"""

import sys
from pathlib import Path
from loguru import logger


def setup_logging(
    log_dir: Path,
    level: str = "DEBUG",
    rotation: str = "00:00",     
    retention: str = "30 days",
    console: bool = True,
) -> None:
    """
    Configure loguru sinks.
    Call once at application startup.
    """
    logger.remove()  # remove default stderr sink

    log_dir.mkdir(parents=True, exist_ok=True)

    # Rotating file — full DEBUG detail
    logger.add(
        log_dir / "rpa_{time:YYYY-MM-DD}.log",
        level="DEBUG",
        rotation=rotation,
        retention=retention,
        format=(
            "{time:YYYY-MM-DD HH:mm:ss.SSS} | "
            "{level:<8} | "
            "{name}:{function}:{line} | "
            "{message}"
        ),
        encoding="utf-8",
        backtrace=True,
        diagnose=True,
    )

    logger.add(
        log_dir / "rpa_errors.log",
        level="ERROR",
        rotation="10 MB",
        retention="90 days",
        format="{time:YYYY-MM-DD HH:mm:ss.SSS} | {level} | {name}:{line} | {message}",
        encoding="utf-8",
    )

    if console:
        logger.add(
            sys.stderr,
            level=level,
            format=(
                "<green>{time:HH:mm:ss}</green> | "
                "<level>{level:<8}</level> | "
                "<cyan>{name}</cyan>:<cyan>{line}</cyan> | "
                "{message}"
            ),
            colorize=True,
        )

    logger.info("Logging initialised. Log dir: {}", log_dir)


__all__ = ["logger", "setup_logging"]
