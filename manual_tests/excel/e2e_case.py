"""
Excel E2E 通用用例驱动器（per-feature 脚本共享）

把 ppt ``run_single_test`` 的「TEST_CASES + validator + ``--test N/all``」范式抽成 Excel
通用骨架，供 8 个 per-feature 子目录（read_state_e2e / range_e2e / ...）复用，避免每个
脚本重抄一份 argparse + 运行循环。

每个 case 用 :class:`ExcelCase` 声明：夹具 + 动作 + 业务参数 + 验证器（成功路径）或
``expect_error_code``（错误码路径）。验证器签名兼容单参 ``(data)`` 与双参 ``(data, reader)``，
后者用 openpyxl :class:`WorkbookReader` 做读盘双重验证。

使用方式（脚本末尾）:
    from manual_tests.excel.e2e_case import ExcelCase, run_main

    TEST_CASES = [ExcelCase(name=..., fixture_name="read_state_e2e/multi_sheet.xlsx", ...)]

    if __name__ == "__main__":
        run_main("Excel Read State E2E", TEST_CASES)
"""

from __future__ import annotations

import argparse
import asyncio
import inspect
import os
import sys
import time
import traceback
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from manual_tests.error_case import (
    PENDING_ADDIN_80,
    PENDING_ADDIN_92,
    ErrorCaseFields,
    ErrorCaseVerdict,
    evaluate_error_case,
    judge_error_case,
    parse_error_details,
)
from manual_tests.excel.e2e_base import ExcelTestRunner, WorkbookReader, _run_applescript
from manual_tests.excel.test_helpers import excel_op

# 判定逻辑已上提 manual_tests/error_case.py（#90，三命名空间共享）。此处再导出以保持
# 既有导入路径可用（含 tests/unit_tests/manual_tests/test_e2e_case_eval.py 与各 per-feature 脚本）。
__all__ = [
    "EXCEL_FIXTURES_ROOT",
    "PENDING_ADDIN_80",
    "PENDING_ADDIN_92",
    "ErrorCaseFields",
    "ErrorCaseVerdict",
    "ExcelCase",
    "case_kind",
    "evaluate_error_case",
    "judge_error_case",
    "parse_error_details",
    "run_cases",
    "run_main",
]

# 所有 per-feature 夹具的根：manual_tests/excel/fixtures/
EXCEL_FIXTURES_ROOT = Path(__file__).parent / "fixtures"

# 验证器：单参 (data) 或双参 (data, reader)，均返回 bool
Validator = Callable[..., bool]


@dataclass
class ExcelCase(ErrorCaseFields):
    """单条 Excel E2E 用例。

    Attributes:
        name: 用例名。
        fixture_name: 相对 ``fixtures/`` 的夹具路径（如 ``"read_state_e2e/multi_sheet.xlsx"``）。
        description: 用例描述。
        action: 不含 ``excel:`` 前缀的动作名（如 ``"get:workbookInfo"``）。
        params: snake_case 业务参数（无需 ``document_uri``）。
        validator: 成功路径验证器（可选）；双参时第二参为 :class:`WorkbookReader`。
        pre_ops: 量测动作前的预备操作 ``[(action, params), ...]``（如先 set 再 get）。
        flow: 自定义多步流——``async (workspace, document_uri, reader) -> bool``。置位时**取代**
            标准「单 action + validator」路径（在 pre_ops 之后运行），用于需要「先 insert 拿
            自动生成的名字，再 update/delete，最后 get 回读核对」这类无法用单 action 表达的场景。
        tags: 标签。

    错误码路径三件套（``expect_error_code`` / ``expect_error_details`` / ``xfail_reason``）
    继承自 :class:`manual_tests.error_case.ErrorCaseFields`，语义见其 docstring。
    """

    name: str
    fixture_name: str
    description: str
    action: str = ""
    params: dict[str, Any] = field(default_factory=dict)
    validator: Validator | None = None
    pre_ops: list[tuple[str, dict[str, Any]]] = field(default_factory=list)
    flow: Callable[..., Any] | None = None
    select_hint: str | None = None
    applescript_select: str | None = None
    tags: list[str] = field(default_factory=list)


def _call_validator(validator: Validator, data: dict[str, Any], reader: WorkbookReader) -> bool:
    """按验证器参数个数智能调用（1 参=仅 data；2 参=双重验证）。"""
    sig = inspect.signature(validator)
    if len(sig.parameters) >= 2:
        return validator(data, reader)
    return validator(data)


async def _run_single(runner: ExcelTestRunner, case: ExcelCase, number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {number}: {case.name}")
    print("=" * 70)
    print(f"📋 描述: {case.description}")
    if case.expect_error_code:
        print(f"🎯 预期错误码: {case.expect_error_code}")

    try:
        async with runner.run_with_workspace(case.fixture_name, open_delay=3.0) as (workspace, fixture):
            for op_action, op_params in case.pre_ops:
                ok, _, err = await excel_op(workspace, fixture.document_uri, op_action, **op_params)
                if not ok:
                    print(f"   ⚠️  预备操作失败 excel:{op_action}: {err}")

            # 选区依赖类（如 get:selectedRange）：优先用 AppleScript 自动选区（全自动、可后台跑）；
            # 设了 applescript_select 就不需要人工。失败时回退到 select_hint 人工提示。
            if case.applescript_select:
                ok_sel, out_sel = _run_applescript(
                    f'tell application "Microsoft Excel"\n'
                    f"    activate\n"
                    f'    select range "{case.applescript_select}"\n'
                    f"end tell"
                )
                print(f"   🖱️  AppleScript 选区 {case.applescript_select!r} → {'ok' if ok_sel else out_sel}")
                await asyncio.sleep(0.5)

            # EXCEL_E2E_PAUSE=1 时人工选区（无 AppleScript 自动化时的回退路径）。
            if case.select_hint and os.environ.get("EXCEL_E2E_PAUSE"):
                print(f"\n👉 请在 Excel 中{case.select_hint}，选好后按回车继续...")
                try:
                    input()
                except EOFError:
                    print("   (无 stdin，跳过暂停；将读取当前选区)")

            # 自定义多步 flow（取代标准单 action 路径）。
            if case.flow is not None:
                reader = WorkbookReader(fixture.working_path)
                print("\n🔀 自定义 flow 执行...")
                passed = await case.flow(workspace, fixture.document_uri, reader)
                _verdict(number, passed)
                return passed

            print(f"\n📝 执行: excel:{case.action} (params={case.params})...")
            start = time.time()
            ok, data, err = await excel_op(workspace, fixture.document_uri, case.action, **case.params)
            print(f"\n⏱️  执行时间: {(time.time() - start) * 1000:.1f}ms")

            # 错误码路径：预期失败 + 校验权威码（+ details 子集）；xfail 三态见 error_case.judge_error_case
            if case.expect_error_code:
                verdict, message = judge_error_case(
                    ok,
                    err,
                    case.expect_error_code,
                    case.expect_error_details,
                    case.xfail_reason,
                )
                print(message)
                _verdict(number, verdict.passed)
                return verdict.passed

            # 成功路径
            if not ok:
                print(f"❌ 调用失败: {err}")
                _verdict(number, False)
                return False

            print("✅ 协议返回成功")
            data = data or {}
            passed = True
            if case.validator:
                await asyncio.sleep(0.5)
                reader = WorkbookReader(fixture.working_path)
                print("\n📊 双重验证:")
                passed = _call_validator(case.validator, data, reader)
            _verdict(number, passed)
            return passed

    except Exception as e:  # noqa: BLE001
        print(f"\n❌ 测试异常: {e}")
        traceback.print_exc()
        return False


def _verdict(number: int, passed: bool) -> None:
    print("\n" + "=" * 70)
    print(f"{'✅' if passed else '❌'} 测试 {number} {'通过' if passed else '失败'}")
    print("=" * 70)


async def run_cases(
    cases: list[ExcelCase],
    indices: list[int],
    *,
    auto_open: bool = True,
    cleanup_on_success: bool = True,
) -> bool:
    """按 indices 顺序运行用例，返回全部是否通过。"""
    # 真机人机配合时给手动激活 Add-In 留足时间：EXCEL_E2E_TIMEOUT（秒），默认 30。
    timeout = float(os.environ.get("EXCEL_E2E_TIMEOUT", "30"))
    runner = ExcelTestRunner(
        fixtures_dir=EXCEL_FIXTURES_ROOT,
        connection_timeout=timeout,
        auto_open=auto_open,
        cleanup_on_success=cleanup_on_success,
    )
    results: list[bool] = []
    for idx in indices:
        if idx < 1 or idx > len(cases):
            continue
        if len(indices) > 1 and results:
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续下一个用例...")
        results.append(await _run_single(runner, cases[idx - 1], idx))
    if len(results) > 1:
        print(f"\n📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    return bool(results) and all(results)


def run_main(title: str, cases: list[ExcelCase]) -> None:
    """脚本入口：解析 ``--test N/all`` / ``--list`` / ``--keep`` 并运行。"""
    parser = argparse.ArgumentParser(description=title)
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(cases) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true", help="不自动打开/激活，逐步人工配合")
    parser.add_argument("--keep", action="store_true", help="成功后保留工作副本供目测（不自动关闭）")
    parser.add_argument("--list", action="store_true", help="仅列出用例")
    args = parser.parse_args()

    if args.list:
        print(f"📋 {title}（{len(cases)} 个用例）:")
        for i, c in enumerate(cases, 1):
            tag = f"  [{case_kind(c)}]"
            print(f"  {i}. {c.name} — {c.description}{tag}")
        return

    indices = list(range(1, len(cases) + 1)) if args.test == "all" else [int(args.test)]
    try:
        success = asyncio.run(
            run_cases(
                cases,
                indices,
                auto_open=not args.no_auto_open,
                cleanup_on_success=not args.keep,
            )
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


def case_kind(case: ExcelCase) -> str:
    """用例类别标签（用于 --list 展示）。"""
    if case.expect_error_code:
        return f"err {case.expect_error_code}" + (" xfail" if case.xfail_reason else "")
    return ",".join(case.tags) if case.tags else "ok"
