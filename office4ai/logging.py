"""
Centralized logging configuration for Office4AI.

Bridges stdlib logging → loguru so that all loggers (including third-party
libraries like socketio, engineio, aiohttp, uvicorn) are unified under a
single format and routed to both console and file sinks.

Environment variables:
    OFFICE4AI_LOG_DIR     – directory for log files (default: platform-standard path, empty string disables file logging)
    OFFICE4AI_LOG_LEVEL   – minimum log level (default: "INFO")
    OFFICE4AI_LOG_CONSOLE – enable console output (default: "true")

Default log directory (when ``OFFICE4AI_LOG_DIR`` is not set):
    - macOS:   ``~/Library/Logs/office4ai``
    - Linux:   ``~/.local/state/office4ai/log``
    - Windows: ``C:\\Users\\<user>\\AppData\\Local\\office4ai\\Logs``
"""

from __future__ import annotations

import logging
import os
import sys

from loguru import logger
from platformdirs import user_log_dir

# Shared format string (no color tags – loguru adds color automatically for console sinks)
_LOG_FORMAT = "{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {name}:{function}:{line} - {message}"

# Third-party loggers to intercept
_INTERCEPTED_LOGGERS = ("socketio", "engineio", "aiohttp", "uvicorn", "uvicorn.access", "uvicorn.error")

# Environment variable for custom log directory
LOG_DIR_ENV = "OFFICE4AI_LOG_DIR"


def get_log_dir() -> str:
    """
    Get the log directory path.

    Priority:
    1. ``OFFICE4AI_LOG_DIR`` environment variable (empty string disables file logging)
    2. ``platformdirs.user_log_dir("office4ai")`` (platform-standard default)

    Returns:
        Path string for the log directory, or empty string to disable file logging.
    """
    env_dir = os.environ.get(LOG_DIR_ENV)
    if env_dir is not None:
        return env_dir
    return user_log_dir("office4ai")


class InterceptHandler(logging.Handler):
    """Route stdlib logging records to loguru, preserving caller information."""

    def emit(self, record: logging.LogRecord) -> None:
        # Map stdlib level to loguru level name
        try:
            level: str | int = logger.level(record.levelname).name
        except ValueError:
            level = record.levelno

        # Find caller frame outside the logging/loguru stack
        frame = sys._getframe(6)
        depth = 6
        while frame and frame.f_code.co_filename == logging.__file__:
            frame = frame.f_back  # type: ignore[assignment]
            depth += 1

        logger.opt(depth=depth, exception=record.exc_info).log(level, record.getMessage())


def setup_logging(
    log_dir: str | None = None,
    log_level: str | None = None,
    console: bool | None = None,
) -> None:
    """
    Configure loguru sinks and bridge stdlib logging.

    Parameters take effect only when the corresponding environment variable
    is **not** set – env vars always win.

    Args:
        log_dir: Directory for log files. ``""`` disables file logging.
        log_level: Minimum log level (e.g. ``"DEBUG"``, ``"INFO"``).
        console: Whether to emit to stderr with colors.
    """
    # --- resolve effective values (env var > argument > default) ---
    env_log_dir = os.environ.get(LOG_DIR_ENV)
    if env_log_dir is not None:
        effective_dir = env_log_dir
    elif log_dir is not None:
        effective_dir = log_dir
    else:
        effective_dir = get_log_dir()
    effective_level = os.environ.get("OFFICE4AI_LOG_LEVEL", log_level if log_level is not None else "INFO").upper()
    console_raw = os.environ.get("OFFICE4AI_LOG_CONSOLE")
    if console_raw is not None:
        effective_console = console_raw.lower() in ("1", "true", "yes")
    else:
        effective_console = console if console is not None else True

    # --- reset loguru ---
    logger.remove()

    # --- console sink ---
    if effective_console:
        logger.add(
            sys.stderr,
            level=effective_level,
            format=_LOG_FORMAT,
            colorize=True,
        )

    # --- file sink ---
    if effective_dir:
        os.makedirs(effective_dir, exist_ok=True)
        logger.add(
            os.path.join(effective_dir, "office4ai-mcp_{time:YYYY-MM-DD}.log"),
            level=effective_level,
            format=_LOG_FORMAT,
            rotation="00:00",
            retention="3 days",
            enqueue=True,
            colorize=False,
        )
        logger.info("Log directory: {}", effective_dir)

    # --- bridge stdlib logging → loguru ---
    logging.basicConfig(handlers=[InterceptHandler()], level=0, force=True)

    for name in _INTERCEPTED_LOGGERS:
        lib_logger = logging.getLogger(name)
        lib_logger.handlers = [InterceptHandler()]
        lib_logger.propagate = False
