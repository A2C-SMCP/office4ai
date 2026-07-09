"""LibreOffice(soffice) 无头调用封装 | Headless LibreOffice invocation.

``.dotx`` / ``.xltx`` 模板与 ``.docx`` / ``.xlsx`` 文档本质只差包内一个 content-type 标记，
但 python-docx / openpyxl 一看是 template 类型就拒绝打开。Word/Excel 点「基于模板新建」=
把该标记落成 document；无 Office 时用 LibreOffice 无头转换做**同等且忠实**的动作
（SDT 控件 / 条件格式 / 结构化表格实测全部幸存）。本模块把「定位 soffice + 隔离 profile +
超时执行」的样板集中起来，供 :mod:`..template` 复用。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from .errors import SofficeConversionError, SofficeNotFoundError

#: 覆盖 soffice 路径的环境变量 | env var to override the soffice path.
SOFFICE_ENV_VAR = "OFFICE4AI_SOFFICE"

#: 各平台常见的 soffice 安装路径（找不到 PATH 里的 soffice 时兜底探测）。
_COMMON_SOFFICE_PATHS: tuple[str, ...] = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    "/usr/bin/soffice",  # Linux
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",  # Apple Silicon Homebrew
    r"C:\Program Files\LibreOffice\program\soffice.exe",  # Windows
)


def find_soffice(explicit: str | None = None) -> str:
    """定位 LibreOffice 可执行文件，按优先级：显式入参 → 环境变量 → PATH → 常见路径。

    Locate the LibreOffice executable. 找不到时抛 :class:`SofficeNotFoundError`。
    """
    candidates: list[str | None] = [explicit, os.environ.get(SOFFICE_ENV_VAR)]
    for cand in candidates:
        if cand and (Path(cand).exists() or shutil.which(cand)):
            return cand
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for path in _COMMON_SOFFICE_PATHS:
        if Path(path).exists():
            return path
    raise SofficeNotFoundError(
        "未找到 LibreOffice(soffice) 可执行文件；请安装 LibreOffice 或设置 "
        f"{SOFFICE_ENV_VAR} 环境变量 | LibreOffice (soffice) not found",
    )


def convert(
    src: str | Path,
    out_dir: str | Path,
    target_ext: str,
    *,
    soffice: str | None = None,
    timeout: float = 120.0,
) -> Path:
    """用无头 LibreOffice 把 ``src`` 转成 ``target_ext`` 格式，输出到 ``out_dir``，返回产物路径。

    Convert ``src`` to ``target_ext`` via headless LibreOffice into ``out_dir``.

    - 使用**独立临时 UserInstallation**，避免与用户已开着的 LibreOffice 抢配置而失败。
    - 失败 / 超时 / 未产出目标文件时抛 :class:`SofficeConversionError`。
    """
    src_path = Path(src)
    out_path = Path(out_dir)
    exe = find_soffice(soffice)
    target_ext = target_ext.lstrip(".")
    produced = out_path / f"{src_path.stem}.{target_ext}"

    with tempfile.TemporaryDirectory() as profile:
        cmd = [
            exe,
            "--headless",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to",
            target_ext,
            "--outdir",
            str(out_path),
            str(src_path),
        ]
        try:
            proc = subprocess.run(cmd, check=False, capture_output=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise SofficeConversionError(
                f"LibreOffice 转换超时({timeout}s) | conversion timed out: {src_path.name}",
            ) from exc
        except OSError as exc:  # exe 不可执行 / 探测后被删 等 | not executable, removed after probe, ...
            raise SofficeConversionError(
                f"LibreOffice 无法执行 | cannot execute {exe!r}: {exc}",
            ) from exc

    if proc.returncode != 0:
        stderr = proc.stderr.decode("utf-8", "replace").strip()
        raise SofficeConversionError(
            f"LibreOffice 转换失败(exit={proc.returncode}) | conversion failed: {src_path.name}\n{stderr}",
        )
    if not produced.exists():
        raise SofficeConversionError(
            f"LibreOffice 未产出目标文件 | expected output missing: {produced}",
        )
    return produced
