"""_soffice 转换错误路径单测（mock subprocess）| Unit tests for soffice conversion error paths."""

from __future__ import annotations

import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest

from office4ai.office.authoring.helpers import _soffice
from office4ai.office.authoring.helpers.errors import SofficeConversionError


@pytest.fixture
def fake_soffice(tmp_path: Path) -> str:
    """一个存在的「假 soffice」文件路径，供 find_soffice 通过、随后 mock subprocess。"""
    exe = tmp_path / "soffice"
    exe.write_text("#!/bin/sh\n")
    return str(exe)


def test_nonzero_exit_raises(fake_soffice: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        _soffice.subprocess,
        "run",
        lambda *a, **k: SimpleNamespace(returncode=1, stderr=b"boom"),
    )
    with pytest.raises(SofficeConversionError, match="转换失败|failed"):
        _soffice.convert(tmp_path / "in.dotx", tmp_path, "docx", soffice=fake_soffice)


def test_timeout_raises(fake_soffice: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def _raise(*_a: object, **_k: object) -> None:
        raise subprocess.TimeoutExpired(cmd="soffice", timeout=1.0)

    monkeypatch.setattr(_soffice.subprocess, "run", _raise)
    with pytest.raises(SofficeConversionError, match="超时|timed out"):
        _soffice.convert(tmp_path / "in.dotx", tmp_path, "docx", soffice=fake_soffice, timeout=1.0)


def test_os_error_raises(fake_soffice: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def _raise(*_a: object, **_k: object) -> None:
        raise PermissionError("denied")

    monkeypatch.setattr(_soffice.subprocess, "run", _raise)
    with pytest.raises(SofficeConversionError, match="无法执行|cannot execute"):
        _soffice.convert(tmp_path / "in.dotx", tmp_path, "docx", soffice=fake_soffice)


def test_missing_output_raises(fake_soffice: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # 退出码 0 但没产出目标文件 -> 仍应报错
    monkeypatch.setattr(
        _soffice.subprocess,
        "run",
        lambda *a, **k: SimpleNamespace(returncode=0, stderr=b""),
    )
    with pytest.raises(SofficeConversionError, match="未产出|missing"):
        _soffice.convert(tmp_path / "in.dotx", tmp_path, "docx", soffice=fake_soffice)
