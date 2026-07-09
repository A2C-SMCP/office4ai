"""authoring 运行时沙箱单元测试 | authoring runtime sandbox unit tests (milestone #4 · S1).

覆盖软沙箱四大守卫 + 返回契约：
- happy path：脚本落盘 → ok / path=work_dir / summary 产物清单 / logs
- stdout/stderr 捕获
- 死循环 → wall-clock 超时杀进程组
- 越权写（work_dir 外）→ 拦截，外部文件不被创建
- import allowlist（subprocess 不在名单）→ 拦截
- os.system → 拦截
- 默认禁网（放开 socket import 后仍禁连接）→ 拦截
- template_uri 不存在 → 运行时前置校验失败
"""

from __future__ import annotations

import os
import tempfile
import uuid
from pathlib import Path

import pytest

import office4ai.office.authoring.runtime as runtime_mod
from office4ai.office.authoring import DEFAULT_ALLOWED_IMPORTS, run_script


async def test_happy_path_creates_file(tmp_path: Path) -> None:
    script = "with open('hello.txt', 'w') as f:\n    f.write('hi from sandbox')\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is True
    assert Path(result.path) == tmp_path
    assert (tmp_path / "hello.txt").read_text() == "hi from sandbox"
    names = [item["name"] for item in result.summary["produced"]]
    assert "hello.txt" in names
    assert result.summary["count"] == 1
    assert result.stderr == ""


async def test_stdout_and_stderr_captured(tmp_path: Path) -> None:
    script = "import sys\nprint('to-stdout')\nprint('to-stderr', file=sys.stderr)\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is True
    assert "to-stdout" in result.logs
    assert "to-stderr" in result.stderr


async def test_infinite_loop_killed_by_timeout(tmp_path: Path) -> None:
    script = "while True:\n    pass\n"
    result = await run_script(script, work_dir=tmp_path, timeout=1.0)

    assert result.ok is False
    assert result.timed_out is True
    assert "timeout" in result.stderr.lower()


async def test_write_outside_work_dir_blocked(tmp_path: Path) -> None:
    evil = Path(tempfile.gettempdir()) / f"office4ai_evil_{uuid.uuid4().hex}.txt"
    assert not evil.exists()
    script = f"open({str(evil)!r}, 'w').write('escape')\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is False
    assert "[sandbox]" in result.stderr
    assert not evil.exists(), "sandbox must block the out-of-bounds write before the file is created"


async def test_disallowed_import_blocked(tmp_path: Path) -> None:
    script = "import subprocess\nsubprocess.run(['echo', 'hi'])\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is False
    assert "not allowed" in result.stderr


async def test_os_system_blocked(tmp_path: Path) -> None:
    marker = Path(tempfile.gettempdir()) / f"office4ai_syscall_{uuid.uuid4().hex}"
    script = f"import os\nos.system('touch {marker}')\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is False
    assert "os.system" in result.stderr
    assert not marker.exists()


async def test_network_blocked_even_when_socket_importable(tmp_path: Path) -> None:
    # 放开 socket import，验证「禁网」发生在连接层而非 import 层
    allowed = (*DEFAULT_ALLOWED_IMPORTS, "socket")
    script = "import socket\nsocket.create_connection(('8.8.8.8', 53), timeout=2)\n"
    result = await run_script(script, work_dir=tmp_path, allowed_imports=allowed)

    assert result.ok is False
    assert "network" in result.stderr.lower()


async def test_network_blocked_udp_sendto(tmp_path: Path) -> None:
    # 回归：UDP/datagram 无需 connect 也不得出网（sendto/sendmsg 被 patch）
    allowed = (*DEFAULT_ALLOWED_IMPORTS, "socket")
    script = "import socket\ns = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)\ns.sendto(b'x', ('8.8.8.8', 53))\n"
    result = await run_script(script, work_dir=tmp_path, allowed_imports=allowed)

    assert result.ok is False
    assert "network" in result.stderr.lower()


async def test_template_uri_missing_fails_fast(tmp_path: Path) -> None:
    missing = tmp_path / "nope.dotx"
    result = await run_script("pass", work_dir=tmp_path, template_uri=str(missing))

    assert result.ok is False
    assert "template_uri not found" in result.stderr
    # 早失败路径的 summary 形状须与成功路径一致（消费者做 summary["produced"] 不 KeyError）
    assert result.summary == {"produced": [], "count": 0, "truncated": False}


async def test_symlink_escape_write_blocked(tmp_path: Path) -> None:
    # work_dir 内放一个指向外部目录的符号链接，通过它写 → realpath 解析后仍在外，须被拦
    outside = tmp_path.parent / f"outside_{uuid.uuid4().hex}"
    outside.mkdir()
    work = tmp_path / "wd"
    work.mkdir()
    (work / "link").symlink_to(outside, target_is_directory=True)
    result = await run_script("open('link/escaped.txt', 'w').write('x')\n", work_dir=work)

    assert result.ok is False
    assert "[sandbox]" in result.stderr
    assert not (outside / "escaped.txt").exists()


async def test_os_rename_outside_work_dir_blocked(tmp_path: Path) -> None:
    dst = tmp_path.parent / f"renamed_{uuid.uuid4().hex}.txt"
    script = f"open('src.txt','w').write('x')\nimport os\nos.rename('src.txt', {str(dst)!r})\n"
    result = await run_script(script, work_dir=tmp_path)

    assert result.ok is False
    assert "[sandbox]" in result.stderr
    assert not dst.exists()


async def test_ctypes_blocked(tmp_path: Path) -> None:
    allowed = (*DEFAULT_ALLOWED_IMPORTS, "ctypes")
    script = "import ctypes\nctypes.CDLL(None)\n"
    result = await run_script(script, work_dir=tmp_path, allowed_imports=allowed)

    assert result.ok is False
    assert "ctypes" in result.stderr


async def test_subprocess_non_allowlisted_executable_blocked(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # 放开 subprocess import，但可执行文件不在白名单（默认无 soffice） → subprocess.Popen audit 拦截
    monkeypatch.setattr(runtime_mod, "detect_soffice", lambda: None)
    allowed = (*DEFAULT_ALLOWED_IMPORTS, "subprocess")
    script = "import subprocess\nsubprocess.run(['/bin/echo', 'hi'])\n"
    result = await run_script(script, work_dir=tmp_path, allowed_imports=allowed)

    assert result.ok is False
    assert "subprocess not allowed" in result.stderr


async def test_subprocess_allowlisted_executable_permitted(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # 用 /bin/echo 冒充「探测到的 soffice」验证白名单放行侧（S2 soffice 接缝）
    echo = "/bin/echo"
    if not Path(echo).exists():
        pytest.skip("/bin/echo not present")
    monkeypatch.setattr(runtime_mod, "detect_soffice", lambda: echo)
    allowed = (*DEFAULT_ALLOWED_IMPORTS, "subprocess")
    script = f"import subprocess\nsubprocess.run([{echo!r}, 'from-sandbox'])\n"
    result = await run_script(script, work_dir=tmp_path, allowed_imports=allowed)

    assert result.ok is True, result.stderr
    assert "from-sandbox" in result.logs


async def test_default_work_dir_is_created() -> None:
    # work_dir 省略 → 运行时建受管临时目录，path 指向它
    result = await run_script("open('a.txt','w').write('x')\n")
    try:
        assert result.ok is True
        wd = Path(result.path)
        assert wd.is_dir()
        assert (wd / "a.txt").exists()
    finally:
        if os.path.isdir(result.path):
            for p in Path(result.path).iterdir():
                p.unlink()
            Path(result.path).rmdir()
