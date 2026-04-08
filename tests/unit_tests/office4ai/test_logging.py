"""Unit tests for the centralized logging module."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from unittest.mock import patch

import pytest
from loguru import logger

from office4ai.logging import InterceptHandler, get_log_dir, setup_logging

# conftest.py sets OFFICE4AI_LOG_DIR="" to prevent accidental file creation.
# Tests that verify file logging must temporarily remove it.


def _patch_no_log_dir_env() -> patch.dict:  # type: ignore[type-arg]
    """Return a context manager that removes OFFICE4AI_LOG_DIR from os.environ."""
    env = os.environ.copy()
    env.pop("OFFICE4AI_LOG_DIR", None)
    return patch.dict(os.environ, env, clear=True)


class TestGetLogDir:
    """Tests for get_log_dir() path resolution."""

    def test_default_uses_platformdirs(self) -> None:
        """Without env var, should return platformdirs user_log_dir path."""
        from platformdirs import user_log_dir

        with _patch_no_log_dir_env():
            result = get_log_dir()
        assert result == user_log_dir("office4ai")

    def test_env_var_overrides(self, tmp_path: Path) -> None:
        """OFFICE4AI_LOG_DIR env var should override platformdirs default."""
        custom = str(tmp_path / "custom_logs")
        with patch.dict(os.environ, {"OFFICE4AI_LOG_DIR": custom}):
            assert get_log_dir() == custom

    def test_empty_env_var_disables(self) -> None:
        """Empty string env var should return empty string (disables file logging)."""
        with patch.dict(os.environ, {"OFFICE4AI_LOG_DIR": ""}):
            assert get_log_dir() == ""

    def test_default_path_is_absolute(self) -> None:
        """Default path (from platformdirs) should be an absolute path."""
        with _patch_no_log_dir_env():
            result = get_log_dir()
        assert os.path.isabs(result)


class TestInterceptHandler:
    """Tests for the stdlib → loguru bridge handler."""

    def test_emit_routes_to_loguru(self, tmp_path: Path) -> None:
        """stdlib logger messages should arrive in loguru via InterceptHandler."""
        log_file = tmp_path / "test.log"

        # Reset loguru and add a simple file sink (no enqueue)
        logger.remove()
        logger.add(str(log_file), format="{message}")

        # Attach InterceptHandler to a stdlib logger
        std_logger = logging.getLogger("test_intercept")
        std_logger.handlers = [InterceptHandler()]
        std_logger.setLevel(logging.DEBUG)

        std_logger.info("hello from stdlib")

        logger.remove()

        contents = log_file.read_text()
        assert "hello from stdlib" in contents


class TestSetupLogging:
    """Tests for setup_logging() configuration function."""

    def test_file_logging_creates_directory(self, tmp_path: Path) -> None:
        log_dir = tmp_path / "newlogs"
        assert not log_dir.exists()

        with _patch_no_log_dir_env():
            setup_logging(log_dir=str(log_dir), console=False)
            logger.info("trigger sink creation")
            logger.remove()

        assert log_dir.exists()
        log_files = list(log_dir.glob("office4ai-mcp_*.log"))
        assert len(log_files) == 1

    def test_file_logging_disabled_with_empty_string(self, tmp_path: Path) -> None:
        with _patch_no_log_dir_env():
            setup_logging(log_dir="", console=False)
            logger.info("should not create files")
            logger.remove()

        # No logs/ directory should appear in tmp_path or cwd
        assert not (tmp_path / "logs").exists()

    def test_env_var_overrides(self, tmp_path: Path) -> None:
        env_dir = tmp_path / "env_logs"
        with patch.dict(os.environ, {"OFFICE4AI_LOG_DIR": str(env_dir)}):
            # Pass a different dir via argument – env var should win
            setup_logging(log_dir=str(tmp_path / "arg_logs"), console=False)
            logger.info("env wins")
            logger.remove()

        assert env_dir.exists()
        assert not (tmp_path / "arg_logs").exists()

    def test_console_disabled(self, capsys: pytest.CaptureFixture[str]) -> None:
        setup_logging(log_dir="", console=False)
        logger.info("invisible message")
        logger.remove()

        captured = capsys.readouterr()
        assert "invisible message" not in captured.err

    def test_log_message_written_to_file(self, tmp_path: Path) -> None:
        with _patch_no_log_dir_env():
            setup_logging(log_dir=str(tmp_path), console=False)
            logger.info("persistent message")
            logger.remove()

        log_files = list(tmp_path.glob("office4ai-mcp_*.log"))
        assert len(log_files) == 1
        contents = log_files[0].read_text()
        assert "persistent message" in contents

    def test_log_level_respected(self, tmp_path: Path) -> None:
        with _patch_no_log_dir_env():
            setup_logging(log_dir=str(tmp_path), log_level="WARNING", console=False)
            logger.debug("debug msg")
            logger.info("info msg")
            logger.warning("warning msg")
            logger.remove()

        log_files = list(tmp_path.glob("office4ai-mcp_*.log"))
        assert len(log_files) == 1
        contents = log_files[0].read_text()
        assert "debug msg" not in contents
        assert "info msg" not in contents
        assert "warning msg" in contents
