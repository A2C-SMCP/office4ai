"""
Excel Read State E2E — get:workbookInfo（工作簿信息）

覆盖 ``excel:get:workbookInfo``（无业务参数）。验证以**返回 data** 为主，openpyxl
``sheet_names`` 做存在性双重验证。

运行方式:
    uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test all
    uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test 2
    uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --list
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.read_state_e2e._fixtures import ensure_fixtures

MULTI = "read_state_e2e/multi_sheet.xlsx"
HIDDEN = "read_state_e2e/hidden_sheet.xlsx"


def _sheet_names(data: dict[str, Any]) -> list[str]:
    return [s.get("name") for s in data.get("sheets", [])]


def validate_multi_sheet(data: dict[str, Any], reader: WorkbookReader) -> bool:
    names = _sheet_names(data)
    print(f"   返回 sheets: {names}")
    if len(names) < 3:
        print(f"   ❌ sheets 数量 {len(names)} < 3")
        return False
    reader.reload()
    disk_names = reader.sheet_names
    missing = [n for n in names if n not in disk_names]
    if missing:
        print(f"   ❌ 返回的 sheet 未落盘核对: {missing}（磁盘 {disk_names}）")
        return False
    print(f"   ✅ {len(names)} 个 sheet，与 openpyxl 磁盘表名一致: {disk_names}")
    return True


def validate_hidden_sheet(data: dict[str, Any], reader: WorkbookReader) -> bool:
    hidden = [s for s in data.get("sheets", []) if s.get("isHidden")]
    print(f"   返回 sheets: {[(s.get('name'), s.get('isHidden')) for s in data.get('sheets', [])]}")
    if not hidden:
        print("   ❌ 未发现 isHidden=true 的工作表")
        return False
    hidden_name = hidden[0].get("name")
    reader.reload()
    state = reader.wb[hidden_name].sheet_state if hidden_name in reader.sheet_names else None
    if state != "hidden":
        print(f"   ❌ openpyxl 核对 sheet_state={state!r}（预期 'hidden'）")
        return False
    print(f"   ✅ 隐藏表 {hidden_name!r}（isHidden=true，openpyxl sheet_state=hidden）")
    return True


def validate_active_sheet(data: dict[str, Any], reader: WorkbookReader) -> bool:
    active = data.get("activeSheet")
    names = _sheet_names(data)
    flagged = [s.get("name") for s in data.get("sheets", []) if s.get("isActive")]
    print(f"   activeSheet={active!r}, isActive 标记={flagged}")
    if active not in names:
        print(f"   ❌ activeSheet {active!r} 不在 sheets {names}")
        return False
    if flagged and active not in flagged:
        print(f"   ❌ activeSheet 与 isActive 标记不一致: {active!r} vs {flagged}")
        return False
    print(f"   ✅ activeSheet={active!r} 合法且与 isActive 一致")
    return True


def validate_file_name(data: dict[str, Any], reader: WorkbookReader) -> bool:
    file_name = data.get("fileName")
    expected = reader.path.name  # 工作副本名（multi_sheet_<时间戳>.xlsx）
    print(f"   fileName={file_name!r}, 工作副本={expected!r}")
    if not file_name or not str(file_name).endswith(".xlsx"):
        print("   ❌ fileName 缺失或非 .xlsx")
        return False
    if not str(file_name).startswith("multi_sheet"):
        print(f"   ⚠️  fileName 未以 'multi_sheet' 开头（实际 {file_name!r}）；Excel 可能返回纯文件名")
    print(f"   ✅ fileName={file_name!r}")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="多 sheet 工作簿",
        fixture_name=MULTI,
        description="3 个工作表，sheets 列表完整且与磁盘一致",
        action="get:workbookInfo",
        validator=validate_multi_sheet,
        tags=["basic"],
    ),
    ExcelCase(
        name="含隐藏 sheet",
        fixture_name=HIDDEN,
        description="返回某 sheet isHidden=true，openpyxl 核对 sheet_state=hidden",
        action="get:workbookInfo",
        validator=validate_hidden_sheet,
        tags=["hidden"],
    ),
    ExcelCase(
        name="activeSheet 正确",
        fixture_name=MULTI,
        description="activeSheet 合法且与 sheets[].isActive 一致",
        action="get:workbookInfo",
        validator=validate_active_sheet,
        tags=["active"],
    ),
    ExcelCase(
        name="fileName 字段",
        fixture_name=MULTI,
        description="fileName 为 .xlsx 文件名",
        action="get:workbookInfo",
        validator=validate_file_name,
        tags=["filename"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Read State E2E — get:workbookInfo", TEST_CASES)
