"""
Excel E2E 测试基础设施

提供 Excel 专属的工作簿读取器（openpyxl 双重验证）、AppleScript 函数（保存 / 关闭 /
激活 Add-In）和测试运行器。结构镜像 ``manual_tests/ppt/e2e_base.py``（PPT 版），仅把
PowerPoint 换成 Excel、python-pptx 换成 openpyxl。

使用方式:
    from manual_tests.excel.e2e_base import ExcelTestRunner, WorkbookReader

    runner = ExcelTestRunner()
    async with runner.run_with_workspace("empty.xlsx") as (workspace, fixture):
        result = await workspace.execute(action)
        reader = WorkbookReader(fixture.working_path)
        reader.reload()           # 先 AppleScript 让 Excel 存盘，再 openpyxl 读
        assert reader.cell_value("Sheet1", "A1") == "Hello Excel"
"""

from __future__ import annotations

import asyncio
import platform
import subprocess
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook

from manual_tests.e2e_base import (
    DocumentFixture,
    E2ETestRunner,
)
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 配置常量
# ==============================================================================

DEFAULT_EXCEL_ADDIN_NAME = "excel-editor4ai"

FIXTURES_ROOT = Path(__file__).parent / "fixtures"

TEMP_ROOT = Path(__file__).parent / ".test_working"


# ==============================================================================
# WorkbookReader —— openpyxl 双重验证
# ==============================================================================


@dataclass
class WorkbookReader:
    """
    Excel 工作簿内容读取器（只读，用于双重验证）。

    通过 openpyxl 读取 Excel 写入磁盘后的 .xlsx，验证单元格值 / 工作表 / 表格 /
    合并区域 / 自动筛选 / 图表数等结构。注意 openpyxl 不渲染、不计算公式，故公式单元格
    读到的是公式字符串或上次缓存值（取决于写入方式）。

    Example:
        reader = WorkbookReader(fixture.working_path)
        reader.reload()
        assert reader.cell_value("Data", "A1") == "Region"
        assert "Sheet3" in reader.sheet_names
    """

    path: Path
    auto_save: bool = True
    _wb: Workbook | None = field(default=None, repr=False)

    @property
    def wb(self) -> Workbook:
        """懒加载 Workbook 对象（data_only=False，保留公式字符串）。"""
        if self._wb is None:
            from openpyxl import load_workbook

            self._wb = load_workbook(str(self.path))
        return self._wb

    def reload(self) -> None:
        """重新加载工作簿；auto_save 时先经 AppleScript 强制 Excel 存盘。"""
        if self.auto_save:
            save_excel_document(self.path)
        self._wb = None

    @property
    def sheet_names(self) -> list[str]:
        """所有工作表名（按 Excel 中的顺序）。"""
        return list(self.wb.sheetnames)

    def cell_value(self, sheet: str, address: str) -> Any:
        """读取单元格值（如 ``cell_value("Sheet1", "A1")``）。"""
        if sheet not in self.wb.sheetnames:
            return None
        return self.wb[sheet][address].value

    def cell_number_format(self, sheet: str, address: str) -> str | None:
        """读取单元格数字格式（如 '0.00'）。"""
        if sheet not in self.wb.sheetnames:
            return None
        return self.wb[sheet][address].number_format

    def cell_fill_hex(self, sheet: str, address: str) -> str | None:
        """读取单元格填充色（'#RRGGBB'），无填充返回 None。"""
        if sheet not in self.wb.sheetnames:
            return None
        fill = self.wb[sheet][address].fill
        rgb = getattr(getattr(fill, "fgColor", None), "rgb", None)
        if not rgb or not isinstance(rgb, str):
            return None
        # openpyxl 通常返回 8 位 ARGB（如 'FFFF0000'）；取后 6 位作为 RRGGBB。
        hex6 = rgb[-6:].upper()
        if hex6 in ("000000", "FFFFFF") and rgb[:2] == "00":
            return None
        return f"#{hex6}"

    def merged_ranges(self, sheet: str) -> list[str]:
        """指定工作表的所有合并区域（如 ['A1:C1']）。"""
        if sheet not in self.wb.sheetnames:
            return []
        return [str(r) for r in self.wb[sheet].merged_cells.ranges]

    def table_names(self, sheet: str) -> list[str]:
        """指定工作表的结构化表格名列表。"""
        if sheet not in self.wb.sheetnames:
            return []
        return list(self.wb[sheet].tables.keys())

    def auto_filter_ref(self, sheet: str) -> str | None:
        """指定工作表的自动筛选范围引用（如 'A1:C10'），无则 None。"""
        if sheet not in self.wb.sheetnames:
            return None
        return self.wb[sheet].auto_filter.ref

    def chart_count(self, sheet: str) -> int:
        """指定工作表的图表数（openpyxl best-effort：读 ws._charts）。"""
        if sheet not in self.wb.sheetnames:
            return 0
        return len(getattr(self.wb[sheet], "_charts", []))


# ==============================================================================
# AppleScript 函数（目标应用：Microsoft Excel）
# ==============================================================================


def _run_applescript(script: str, timeout: float = 15.0) -> tuple[bool, str]:
    """执行 AppleScript 并返回 (是否成功, 输出/错误)。"""
    try:
        result = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=timeout)
        if result.returncode == 0:
            return True, result.stdout.decode().strip()
        return False, result.stderr.decode().strip()
    except subprocess.TimeoutExpired:
        return False, "AppleScript 执行超时"
    except Exception as e:  # noqa: BLE001
        return False, str(e)


def save_excel_document(path: Path) -> bool:
    """强制 Excel 保存工作簿（通过 AppleScript）。在 openpyxl 读盘验证前调用。"""
    if platform.system() != "Darwin":
        print("⚠️  自动保存仅支持 macOS")
        return False

    doc_name = path.name
    doc_stem = path.stem
    script = f'''
    tell application "Microsoft Excel"
        set targetBooks to (every workbook whose name is "{doc_name}")
        if (count of targetBooks) is 0 then
            set targetBooks to (every workbook whose name starts with "{doc_stem}")
        end if
        repeat with b in targetBooks
            save b
        end repeat
        return count of targetBooks
    end tell
    '''
    ok, output = _run_applescript(script, timeout=8.0)
    if not ok:
        print(f"⚠️  AppleScript 保存错误: {output}")
        return False
    if output == "0":
        print("⚠️  未找到匹配的工作簿进行保存")
        return False
    return True


def close_excel_document(path: Path) -> bool:
    """关闭 Excel 工作簿（不保存，通过 AppleScript）。"""
    if platform.system() != "Darwin":
        return False

    doc_name = path.name
    doc_stem = path.stem
    script = f'''
    tell application "Microsoft Excel"
        set targetBooks to (every workbook whose name is "{doc_name}")
        if (count of targetBooks) is 0 then
            set targetBooks to (every workbook whose name starts with "{doc_stem}")
        end if
        repeat with b in targetBooks
            close b saving no
        end repeat
        return count of targetBooks
    end tell
    '''
    ok, output = _run_applescript(script, timeout=8.0)
    if ok:
        print(f"📕 已关闭工作簿: {doc_name} ({output} matched)")
        return True
    print(f"⚠️  AppleScript 关闭错误: {output}")
    return False


async def activate_excel_addin(addin_name: str = DEFAULT_EXCEL_ADDIN_NAME) -> bool:
    """通过 AppleScript UI 自动化尝试激活 Excel Add-In（仅 macOS，best-effort）。

    与 PPT 版同构：尝试点击功能区「加载项 / Add-in」按钮；点不到则提示用户手动激活。
    返回 True 仅表示已点击「加载项」按钮，最终激活仍可能需要用户在面板中点选。
    """
    if platform.system() != "Darwin":
        print("⚠️  自动激活 Add-In 仅支持 macOS，请手动激活")
        return False

    print(f"🔌 尝试自动激活 Excel Add-In: {addin_name}...")
    script = """
tell application "Microsoft Excel" to activate
delay 0.5

tell application "System Events"
    tell process "Microsoft Excel"
        set tg to tab group 1 of front window
        set allElems to entire contents of tg
        repeat with elem in allElems
            try
                if role of elem is "AXButton" then
                    set eName to name of elem
                    if eName contains "加载项" or eName contains "Add-in" then
                        click elem
                        return "ok"
                    end if
                end if
            end try
        end repeat
        return "not_found"
    end tell
end tell
"""
    ok, output = _run_applescript(script)
    if not ok or output != "ok":
        print(f"   ⚠️  未能自动点击「加载项」按钮 (原因: {output})")
        print(f"   👆 请手动点击「加载项」→「{addin_name}...」")
        return False
    print("   ✅ 已点击「加载项」按钮")
    print(f"   👆 请在弹出的加载项面板中点击「{addin_name}...」")
    return False


# ==============================================================================
# ExcelTestRunner
# ==============================================================================


class ExcelTestRunner(E2ETestRunner):
    """Excel E2E 运行器：继承通用 E2ETestRunner，覆盖 Add-In 激活与文档关闭逻辑。"""

    @asynccontextmanager
    async def run_with_workspace(
        self,
        fixture_name: str,
        open_delay: float = 3.0,
    ) -> AsyncIterator[tuple[OfficeWorkspace, DocumentFixture]]:
        """准备 Excel 工作簿并启动 Workspace（覆盖父类，使用 Excel 专属激活/关闭）。"""
        async with self.prepare_document(fixture_name, open_delay) as fixture:  # pyright: ignore[reportGeneralTypeIssues]
            workspace = OfficeWorkspace(host=self.host, port=self.port)
            try:
                await workspace.start()
                print("✅ Workspace 启动成功")

                if self.auto_open and self.auto_activate:
                    activated = await activate_excel_addin()
                    if not activated:
                        print(f"👆 请手动点击「加载项」→「{DEFAULT_EXCEL_ADDIN_NAME}...」激活 Add-In")

                print("⏳ 等待 Excel Add-In 连接...")
                connected = await workspace.wait_for_addin_connection(timeout=self.connection_timeout)
                if not connected:
                    raise RuntimeError("超时：未检测到 Excel Add-In 连接")
                print("✅ Excel Add-In 已连接")

                yield workspace, fixture
            finally:
                if self.auto_close and not fixture.document_closed and fixture.cleanup_on_success:
                    close_excel_document(fixture.working_path)
                    fixture.document_closed = True
                    await asyncio.sleep(0.5)
                await workspace.stop()


# ==============================================================================
# 夹具工具
# ==============================================================================


def create_empty_xlsx(path: Path) -> None:
    """生成一个仅含单个空 'Sheet1' 的 .xlsx 夹具。"""
    from openpyxl import Workbook as _Workbook

    wb = _Workbook()
    ws = wb.active
    if ws is not None:
        ws.title = "Sheet1"
    wb.save(str(path))


def ensure_empty_fixture() -> Path:
    """确保 empty.xlsx 夹具存在，返回其路径。"""
    FIXTURES_ROOT.mkdir(parents=True, exist_ok=True)
    empty_path = FIXTURES_ROOT / "empty.xlsx"
    if not empty_path.exists():
        print(f"📝 创建夹具: {empty_path}")
        create_empty_xlsx(empty_path)
    return empty_path
