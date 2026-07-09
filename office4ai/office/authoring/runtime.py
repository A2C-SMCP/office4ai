"""authoring 运行时 | authoring runtime (parent side).

``office_run_script`` 的服务端运行时：把 LLM 提交的 Python 脚本放进一个跨平台
软沙箱子进程里执行，并回传结构化结果。

设计立场（milestone #4 · S1）：
    - **脚本全自主**：操作哪个文件、是否保存、任何动作都在脚本里；运行时不注入
      产物路径、不替脚本决定写到哪。运行时只提供「在哪跑（work_dir）+ 白名单边界 +
      稳定依赖」，并观察脚本产出了什么。
    - **返回契约**（locked）：``{ok, path, summary, logs, stderr}``
        * ``ok``     —— 脚本是否运行成功（exit 0、无未捕获异常、未被超时/沙箱拦截）
        * ``path``   —— 脚本运行目录（work_dir，也是 FS 写白名单），调用方据此找产物
        * ``summary``—— 运行时事后扫描 work_dir 得到的**产物清单**（新增/改动文件）；
                        不解析 office 文件内部结构（pages/slides/sheets 的富视觉呈现留给
                        W4b per-file window）
        * ``logs``   —— 捕获的 stdout
        * ``stderr`` —— 捕获的 stderr（沙箱拦截原因、脚本 traceback、超时说明都在此）

沙箱语义（D4，软沙箱/跨平台/best-effort）：子进程 + 超时杀进程组（死循环）+
FS 写白名单（越权写）+ import allowlist + 默认禁网。强隔离（Linux bubblewrap/
nsjail）留作后续 OS 级加固接缝。
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import signal
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from loguru import logger

from office4ai.environment.workspace.services.document_lock import document_lock_manager

# ---------------------------------------------------------------------------
# 常量 | Constants
# ---------------------------------------------------------------------------

#: 默认 wall-clock 超时（秒）—— 对齐 OASP「批量操作 60s」约定
DEFAULT_TIMEOUT_SECONDS: float = 60.0
#: 产物清单最多列出的条目数（防脚本产生海量文件撑爆返回）
_MANIFEST_CAP: int = 500

#: 用户脚本可直接 import 的顶层名单（库的传递 import 不受此限）。
#: office4ai.office.authoring.* 另在子进程侧特判放行（S2 helper 随 SKILL 分发）。
DEFAULT_ALLOWED_IMPORTS: tuple[str, ...] = (
    # office / 模板库
    "docx",
    "pptx",
    "openpyxl",
    "lxml",
    "docxtpl",
    "jinja2",
    "PIL",
    # 安全 stdlib 子集
    "os",
    "sys",
    "io",
    "re",
    "json",
    "csv",
    "math",
    "cmath",
    "random",
    "datetime",
    "time",
    "pathlib",
    "collections",
    "itertools",
    "functools",
    "operator",
    "typing",
    "dataclasses",
    "string",
    "textwrap",
    "decimal",
    "fractions",
    "statistics",
    "uuid",
    "base64",
    "hashlib",
    "hmac",
    "secrets",
    "zipfile",
    "gzip",
    "tempfile",
    "shutil",
    "glob",
    "copy",
    "enum",
    "abc",
    "contextlib",
    "warnings",
    "traceback",
    "struct",
    "binascii",
    "unicodedata",
    "html",
    "xml",
)

_CHILD_PATH = Path(__file__).with_name("_sandbox_child.py")

# soffice 探测候选（含 macOS app bundle 内路径）
_SOFFICE_CANDIDATES: tuple[str, ...] = (
    "soffice",
    "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
)


# ---------------------------------------------------------------------------
# 返回契约 | Result contract
# ---------------------------------------------------------------------------


def _empty_manifest() -> dict[str, Any]:
    """空产物清单 —— 保证 ``summary`` 在所有返回路径（含早失败）形状一致。"""
    return {"produced": [], "count": 0, "truncated": False}


@dataclass
class ScriptResult:
    """脚本运行结果 | Result of a sandboxed script run."""

    ok: bool
    path: str
    summary: dict[str, Any] = field(default_factory=_empty_manifest)
    logs: str = ""
    stderr: str = ""
    # —— 内部诊断字段，不进 5 键契约 | internal diagnostics (not in the 5-key wire contract) ——
    returncode: int | None = None
    timed_out: bool = False

    def to_contract(self) -> dict[str, Any]:
        """收敛为 locked 的 5 键返回契约 | Collapse to the locked 5-key contract."""
        return {
            "ok": self.ok,
            "path": self.path,
            "summary": self.summary,
            "logs": self.logs,
            "stderr": self.stderr,
        }


# ---------------------------------------------------------------------------
# soffice 探测 | soffice detection
# ---------------------------------------------------------------------------

_UNSET: object = object()
_soffice_cache: str | None | object = _UNSET


def detect_soffice() -> str | None:
    """探测 LibreOffice ``soffice`` 可执行文件路径，找不到返回 ``None``（缓存结果）。

    S2 的 ``instantiate_from_template()``（dotx/xltx 实例化）依赖它；沙箱会把探测到的
    路径加入 subprocess 白名单，使 helper 能在沙箱内调 soffice。
    """
    global _soffice_cache
    if _soffice_cache is not _UNSET:
        return _soffice_cache  # type: ignore[return-value]
    found: str | None = None
    for cand in _SOFFICE_CANDIDATES:
        if os.sep in cand:
            resolved = cand if os.path.exists(cand) else None
        else:
            resolved = shutil.which(cand)
        if resolved:
            found = str(Path(resolved).resolve())
            break
    _soffice_cache = found
    return found


def log_soffice_status() -> None:
    """启动期日志：报告 soffice 是否可用（缺失仅告警，不阻断）。"""
    path = detect_soffice()
    if path:
        logger.info(f"authoring: LibreOffice soffice 已探测到 | soffice detected at {path}")
    else:
        logger.warning(
            "authoring: 未探测到 LibreOffice soffice —— 模板实例化（.dotx/.xltx）将不可用 | "
            "soffice not found; template instantiation (.dotx/.xltx) unavailable",
        )


# ---------------------------------------------------------------------------
# 产物清单 | Product manifest
# ---------------------------------------------------------------------------


def _snapshot(root: Path) -> dict[str, tuple[int, float]]:
    """浅递归快照 work_dir 下的文件 (relpath -> (size, mtime))。"""
    snap: dict[str, tuple[int, float]] = {}
    if not root.exists():
        return snap
    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            fp = Path(dirpath) / name
            try:
                st = fp.stat()
            except OSError:
                continue
            snap[str(fp.relative_to(root))] = (st.st_size, st.st_mtime)
    return snap


def _diff_manifest(before: dict[str, tuple[int, float]], after: dict[str, tuple[int, float]]) -> dict[str, Any]:
    """由前后快照算出新增/改动文件清单。"""
    produced: list[dict] = []
    for rel, (size, mtime) in sorted(after.items()):
        prev = before.get(rel)
        if prev is None:
            produced.append({"name": rel, "size": size, "status": "created"})
        elif prev != (size, mtime):
            produced.append({"name": rel, "size": size, "status": "modified"})
    truncated = len(produced) > _MANIFEST_CAP
    return {
        "produced": produced[:_MANIFEST_CAP],
        "count": len(produced),
        "truncated": truncated,
    }


# ---------------------------------------------------------------------------
# 运行 | Run
# ---------------------------------------------------------------------------


def _kill_process_group(proc: asyncio.subprocess.Process) -> None:
    """SIGKILL 整个进程组（含脚本可能 fork 出的 soffice 子进程）。"""
    if proc.returncode is not None:
        return
    try:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
    except (ProcessLookupError, PermissionError, OSError):
        try:
            proc.kill()
        except ProcessLookupError:
            pass


#: 传给沙箱子进程的 env 最小白名单 —— 杜绝把服务端机密（云凭证 / API key 等）
#: 透传给不可信脚本（脚本 os 在白名单内、读操作不受限，可读 environ 外泄）。
#: 保留 soffice headless 实例化所需的 HOME/PATH/LANG 等。
_ENV_ALLOWLIST: tuple[str, ...] = (
    "PATH",
    "HOME",
    "USER",
    "LOGNAME",
    "LANG",
    "LC_ALL",
    "LC_CTYPE",
    "TZ",
    "TMPDIR",
    "TEMP",
    "TMP",
    "SYSTEMROOT",  # Windows soffice/py 运行所需
)


def _child_env() -> dict[str, str]:
    env = {k: os.environ[k] for k in _ENV_ALLOWLIST if k in os.environ}
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    return env


async def run_script(
    script: str,
    *,
    work_dir: str | os.PathLike[str] | None = None,
    template_uri: str | None = None,
    timeout: float | None = None,
    allowed_imports: tuple[str, ...] = DEFAULT_ALLOWED_IMPORTS,
) -> ScriptResult:
    """在软沙箱子进程中执行 ``script``，回传 :class:`ScriptResult`。

    Args:
        script: LLM 提交的 Python 源码（内联字符串）。
        work_dir: 脚本运行目录（cwd + FS 写白名单）。缺省创建一个受管临时目录（用后不删，
            由调用方通过 ``path`` 取产物；调用方负责清理或长期保留）。
        template_uri: 可选模板文件路径/URI —— **只读参考**。读操作不受白名单限制，脚本可
            自由读取；本参数仅做存在性前置校验（fail-fast）并向脚本表意。按设计取舍，模板
            **不原地改写**（写白名单只含 work_dir）：脚本应把模板实例化/产物写进 work_dir。
        timeout: wall-clock 超时（秒），缺省 :data:`DEFAULT_TIMEOUT_SECONDS`。
        allowed_imports: 用户脚本可直接 import 的顶层名单。
    """
    timeout = DEFAULT_TIMEOUT_SECONDS if timeout is None else float(timeout)

    # work_dir：缺省建受管临时目录（不自动删——产物要留给调用方）
    if work_dir is None:
        wd = Path(tempfile.mkdtemp(prefix="office4ai-run-"))
    else:
        wd = Path(work_dir).expanduser().resolve()
        wd.mkdir(parents=True, exist_ok=True)

    # 模板存在性校验（best-effort）
    if template_uri:
        tpl_path = _uri_to_path(template_uri)
        if tpl_path is not None and not tpl_path.exists():
            return ScriptResult(
                ok=False,
                path=str(wd),
                stderr=f"[runtime] template_uri not found: {template_uri!r}",
            )

    soffice = detect_soffice()
    config = {
        "script": script,
        "work_dir": str(wd),
        "write_allow": [str(wd)],
        "allowed_imports": list(allowed_imports),
        "allowed_executables": [soffice] if soffice else [],
        "block_network": True,
        "cpu_limit_seconds": int(timeout) + 5,
    }

    # document_lock 串行化对同一 work_dir 的落盘写，防并发 run 相互覆盖
    work_uri = wd.as_uri()
    async with document_lock_manager.acquire(work_uri):
        before = _snapshot(wd)
        result = await _spawn_and_wait(config, wd, timeout)
        after = _snapshot(wd)

    result.summary = _diff_manifest(before, after)
    return result


async def _spawn_and_wait(config: dict[str, Any], wd: Path, timeout: float) -> ScriptResult:
    tmp_cfg_dir = Path(tempfile.mkdtemp(prefix="office4ai-cfg-"))
    cfg_path = tmp_cfg_dir / "config.json"
    cfg_path.write_text(json.dumps(config), encoding="utf-8")
    try:
        proc = await asyncio.create_subprocess_exec(
            sys.executable,
            "-I",
            "-B",
            str(_CHILD_PATH),
            str(cfg_path),
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            cwd=str(wd),
            env=_child_env(),
            start_new_session=True,  # 新会话/进程组，便于超时整组 SIGKILL
        )
        comm = asyncio.ensure_future(proc.communicate())
        timed_out = False
        try:
            stdout_b, stderr_b = await asyncio.wait_for(asyncio.shield(comm), timeout)
        except asyncio.TimeoutError:
            timed_out = True
            _kill_process_group(proc)
            # SIGKILL 后应迅速返回；但若某孙进程卡在不可中断态并持有管道 fd，
            # 再包一层有界超时，宁可丢日志也不让工具无限挂起。
            try:
                stdout_b, stderr_b = await asyncio.wait_for(comm, timeout=5.0)
            except asyncio.TimeoutError:
                comm.cancel()
                stdout_b, stderr_b = b"", b""

        logs = stdout_b.decode("utf-8", errors="replace")
        stderr = stderr_b.decode("utf-8", errors="replace")
        if timed_out:
            stderr = (stderr + f"\n[sandbox] killed: wall-clock timeout {timeout:g}s exceeded").strip()

        ok = (not timed_out) and proc.returncode == 0
        return ScriptResult(
            ok=ok,
            path=str(wd),
            logs=logs,
            stderr=stderr,
            returncode=proc.returncode,
            timed_out=timed_out,
        )
    finally:
        shutil.rmtree(tmp_cfg_dir, ignore_errors=True)


def _uri_to_path(uri: str) -> Path | None:
    """把 ``file://`` URI 或普通路径转成 Path；无法识别返回 None。"""
    if uri.startswith("file://"):
        from urllib.parse import unquote, urlparse

        parsed = urlparse(uri)
        return Path(unquote(parsed.path))
    if "://" in uri:
        return None
    return Path(uri).expanduser()
