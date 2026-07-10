"""authoring 沙箱子进程引导 | authoring sandbox child bootstrap.

本模块是 ``office_run_script`` 沙箱的**子进程侧**入口，被父进程以::

    python -I -B _sandbox_child.py <config.json>

方式拉起（``-I`` 隔离模式忽略用户 site / 环境变量；``-B`` 禁写 .pyc 避免
字节码写入触发写白名单守卫）。**只依赖标准库**，绝不 import ``office4ai``，
以保持沙箱边界清晰、启动开销最小。

执行顺序（关键）：
    1. 读取 config（此时尚未装 audit hook，父进程写的 config 文件可自由读）；
    2. 设置 RLIMIT（内存/CPU/文件大小，作为父进程 wall-clock 杀进程组之外的兜底）；
    3. chdir 进 work_dir；
    4. 预热 import 受信重库（``openpyxl → numpy`` 等，其 import 期的 ``ctypes.dlopen`` 须在装 hook 前完成）；
    5. 装 audit hook（FS 写越权 / 进程逃逸 / ctypes 拦截）——**装之后本进程所有操作都受管**；
    6. monkeypatch socket egress 实现「默认禁网」；
    7. 构造受限 builtins（用户帧 ``__import__`` allowlist）；
    8. ``exec`` 用户脚本，异常打到 stderr 并以非零码退出。

软沙箱定位：拦截「误操作 / 明显越权」（死循环、越权写、联网、逃逸），
而非对抗性 RCE。留 OS 级加固接缝（Linux bubblewrap/nsjail）由后续演进。

**覆盖率说明**：本模块跑在子进程，行级 ``--cov`` 探针不可见（报 0%）。其守卫逻辑
以**行为测试**保证——``tests/.../office/authoring/test_runtime.py`` 与集成测试经真实
子进程逐条驱动（越权写 / symlink 逃逸 / import gate / os.system / subprocess 白名单 /
ctypes / 禁网 / 超时），而非行覆盖。
"""

from __future__ import annotations

import builtins
import json
import os
import sys
import traceback
from typing import Any


class SandboxViolation(PermissionError):
    """沙箱策略拦截 | Raised when the sandboxed script violates a policy."""


# ---------------------------------------------------------------------------
# 路径归属判定 | Path containment checks
# ---------------------------------------------------------------------------


def _real(path: str) -> str:
    """规范化为绝对真实路径（解析符号链接，防 symlink 逃逸）。

    对尚不存在的文件（新建写入）逐级解析已存在的部分即可。
    ``os.path.realpath`` 内部走 ``lstat``（非受审计事件），不会触发 audit
    hook 递归。
    """
    try:
        return os.path.realpath(os.path.abspath(path))
    except (OSError, ValueError):
        return os.path.abspath(path)


def _under(path: str, roots: list[str]) -> bool:
    """path 是否位于任一白名单根目录之下（含根目录本身）。"""
    rp = _real(path)
    for root in roots:
        if rp == root or rp.startswith(root + os.sep):
            return True
    return False


def _is_write_mode(mode: Any, flags: Any) -> bool:
    """由 open 的 mode/flags 判定是否为写打开。

    ``open`` 审计事件对内建 ``open`` 给出 str mode，对 ``os.open`` 给出 int flags。
    """
    if isinstance(mode, str) and any(c in mode for c in ("w", "a", "x", "+")):
        return True
    if isinstance(flags, int) and flags:
        # O_WRONLY=1, O_RDWR=2, 以及 O_CREAT/O_APPEND/O_TRUNC 都视为写意图
        write_bits = os.O_WRONLY | os.O_RDWR | os.O_APPEND | os.O_CREAT | os.O_TRUNC
        if flags & write_bits:
            return True
    return False


# ---------------------------------------------------------------------------
# audit hook（进程级、不可卸载）| audit hook (process-wide, non-removable)
# ---------------------------------------------------------------------------


def _install_audit_hook(write_allow: list[str], allowed_execs: list[str]) -> None:
    exec_reals = {_real(e) for e in allowed_execs}

    def _exec_allowed(executable: Any) -> bool:
        if executable is None:
            return False
        try:
            return _real(os.fspath(executable)) in exec_reals
        except (TypeError, ValueError):
            return False

    def _hook(event: str, args: tuple) -> None:
        # ---- 文件系统写越权 | filesystem write escape ----
        if event == "open":
            path, mode, flags = args
            if path is not None and _is_write_mode(mode, flags) and not _under(os.fspath(path), write_allow):
                raise SandboxViolation(f"[sandbox] blocked write outside work_dir: {path!r}")
        elif event in ("os.rename", "os.replace", "os.link", "os.symlink"):
            # (src, dst) —— 目标端必须在白名单内
            dst = args[1]
            if not _under(os.fspath(dst), write_allow):
                raise SandboxViolation(f"[sandbox] blocked {event} outside work_dir: {dst!r}")
        elif event in ("os.remove", "os.mkdir", "os.rmdir", "os.truncate", "os.chmod", "os.chown", "os.utime"):
            target = args[0]
            if not _under(os.fspath(target), write_allow):
                raise SandboxViolation(f"[sandbox] blocked {event} outside work_dir: {target!r}")
        # ---- 进程逃逸 | process escape ----
        elif event == "os.system":
            raise SandboxViolation("[sandbox] os.system is not allowed")
        elif event == "subprocess.Popen":
            # args: (executable, args, cwd, env)
            if not _exec_allowed(args[0]):
                raise SandboxViolation(f"[sandbox] subprocess not allowed: {args[0]!r}")
        elif event in ("os.exec", "os.posix_spawn"):
            if not _exec_allowed(args[0]):
                raise SandboxViolation(f"[sandbox] exec not allowed: {args[0]!r}")
        elif event.startswith("ctypes."):
            raise SandboxViolation(f"[sandbox] ctypes is not allowed ({event})")
        # ---- 网络 | network（兜底，socket egress 另有 monkeypatch）----
        elif event in ("socket.getaddrinfo", "socket.gethostbyname", "socket.connect"):
            raise SandboxViolation(f"[sandbox] network access is blocked ({event})")

    sys.addaudithook(_hook)


# ---------------------------------------------------------------------------
# 网络 monkeypatch（比 audit 更可靠地覆盖直连 IP 场景）
# ---------------------------------------------------------------------------


def _block_network() -> None:
    try:
        import socket
    except ImportError:  # pragma: no cover - socket 总是可用
        return

    def _denied(*_a: Any, **_k: Any) -> Any:
        raise SandboxViolation("[sandbox] network access is blocked (default deny-net)")

    # 连接层 + 名字解析层
    socket.socket.connect = _denied  # type: ignore[method-assign]
    socket.socket.connect_ex = _denied  # type: ignore[method-assign]
    socket.getaddrinfo = _denied
    socket.gethostbyname = _denied
    if hasattr(socket, "create_connection"):
        socket.create_connection = _denied
    # 发送层 —— 堵住无需 connect 的 UDP/datagram 出网（sendto/sendmsg）
    socket.socket.send = _denied  # type: ignore[method-assign]
    socket.socket.sendall = _denied  # type: ignore[method-assign]
    socket.socket.sendto = _denied  # type: ignore[method-assign]
    socket.socket.sendmsg = _denied  # type: ignore[method-assign]


# ---------------------------------------------------------------------------
# import allowlist（仅约束用户帧的直接 import；库内传递 import 不受影响）
# ---------------------------------------------------------------------------


def _guarded_import(allowed: set[str]) -> Any:
    real_import = builtins.__import__

    def _import(name: str, g: Any = None, ln: Any = None, fromlist: Any = (), level: int = 0) -> Any:
        if level == 0 and name:
            root = name.split(".")[0]
            if root == "office4ai":
                # 仅放行 authoring helper 子包（S2 随 SKILL 分发），杜绝 import 运行中的 server 内部
                if not name.startswith("office4ai.office.authoring"):
                    raise ImportError("[sandbox] only office4ai.office.authoring.* may be imported")
            elif root not in allowed:
                raise ImportError(f"[sandbox] import of {name!r} is not allowed (allowlist gated)")
        return real_import(name, g, ln, fromlist, level)

    return _import


def _set_rlimits(mem_bytes: int, cpu_seconds: int, fsize_bytes: int) -> None:
    """POSIX 资源上限兜底 | Best-effort POSIX resource caps (guarded per-limit)."""
    try:
        import resource
    except ImportError:  # pragma: no cover - 非 POSIX
        return
    for name, soft in (
        ("RLIMIT_CPU", cpu_seconds),
        ("RLIMIT_AS", mem_bytes),
        ("RLIMIT_FSIZE", fsize_bytes),
    ):
        limit = getattr(resource, name, None)
        if limit is None:
            continue
        try:
            hard = resource.getrlimit(limit)[1]
            new_soft = soft if hard == resource.RLIM_INFINITY else min(soft, hard)
            resource.setrlimit(limit, (new_soft, hard))
        except (ValueError, OSError):
            # macOS 对 RLIMIT_AS 常不支持 —— 忽略单项失败
            continue


def _warmup_trusted_libs(names: list[str]) -> None:
    """在装 audit hook 前预热 import 受信重库（best-effort）。

    这些库（如 ``openpyxl → numpy``）在 import 期会触发 ``ctypes.dlopen`` 等被沙箱拦截的
    操作——若发生在 hook 之后会抛 :class:`SandboxViolation`。预热在受信阶段完成其模块
    初始化（含 ``ctypes`` 模块自身首次 import 的 dlopen），之后用户脚本 ``import`` 命中
    ``sys.modules`` 缓存不再触发。``names`` 由父进程按受信库白名单下发（``allowed_imports``
    与原生库集合的交集），用户脚本无法影响。

    **安全语义**（软沙箱、非对抗性 RCE）：预热把 ctypes 等送进 ``sys.modules`` **不扩大
    对抗面**——本沙箱的 import allowlist 与 audit hook 均是 defense-in-depth 而非硬边界
    （真实隔离留给 OS 级加固接缝）。用户帧 ``import ctypes`` 仍被 allowlist 挡；用户对
    ctypes 的**会发 audit 事件**的危险操作（``dlopen`` / ``dlsym`` / ``call_function`` 等）
    仍被进程级 hook 拦，无论经何种途径取到 ctypes 模块对象。

    仅在运行时用 ``importlib`` 动态 import 第三方库（本模块的静态依赖仍只有标准库）；
    单库失败打一行 stderr 诊断并跳过——对预热的原生库，失败会让它在 hook 装好后 import
    时触到 ctypes 拦截成误导性的 SandboxViolation，故此诊断有助排查。
    """
    import importlib

    for name in names:
        try:
            importlib.import_module(name)
        except Exception as exc:  # noqa: BLE001 - best-effort 预热；失败打诊断后继续
            print(f"[sandbox] warmup import {name!r} failed: {type(exc).__name__}: {exc}", file=sys.stderr)
            continue


def main() -> int:
    with open(sys.argv[1], encoding="utf-8") as fh:
        cfg = json.load(fh)

    script: str = cfg["script"]
    work_dir: str = cfg["work_dir"]
    write_allow: list[str] = [_real(p) for p in cfg["write_allow"]]
    allowed_imports: set[str] = set(cfg["allowed_imports"])
    allowed_execs: list[str] = cfg.get("allowed_executables", [])
    block_net: bool = cfg.get("block_network", True)

    _set_rlimits(
        mem_bytes=int(cfg.get("mem_limit_bytes", 2 * 1024**3)),
        cpu_seconds=int(cfg.get("cpu_limit_seconds", 30)),
        fsize_bytes=int(cfg.get("fsize_limit_bytes", 512 * 1024**2)),
    )

    os.chdir(work_dir)

    # 预热受信重库（须在装 hook 前）：openpyxl→numpy 等 import 期会触发 ctypes.dlopen，
    # 装 hook 后会被拦成 SandboxViolation；预热后用户脚本 import 命中缓存不再触发。
    _warmup_trusted_libs(cfg.get("warmup_imports", []))

    # —— 自此以下装上守卫，用户脚本全程受管 ——
    _install_audit_hook(write_allow, allowed_execs)
    if block_net:
        _block_network()

    sandbox_builtins = dict(vars(builtins))
    sandbox_builtins["__import__"] = _guarded_import(allowed_imports)
    script_globals: dict[str, Any] = {
        "__name__": "__office_run_script__",
        "__builtins__": sandbox_builtins,
        "__file__": "<office_run_script>",
    }

    try:
        code = compile(script, "<office_run_script>", "exec")
        exec(code, script_globals)  # noqa: S102 - 沙箱化执行用户脚本，边界即本模块
    except BaseException:  # noqa: BLE001 - 捕获一切（含 SystemExit 之外）汇报给父进程
        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
