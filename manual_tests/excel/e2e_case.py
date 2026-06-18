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
import sys
import time
import traceback
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from manual_tests.excel.e2e_base import ExcelTestRunner, WorkbookReader
from manual_tests.excel.test_helpers import excel_op

# 所有 per-feature 夹具的根：manual_tests/excel/fixtures/
EXCEL_FIXTURES_ROOT = Path(__file__).parent / "fixtures"

# 验证器：单参 (data) 或双参 (data, reader)，均返回 bool
Validator = Callable[..., bool]


@dataclass
class ExcelCase:
    """单条 Excel E2E 用例。

    Attributes:
        name: 用例名。
        fixture_name: 相对 ``fixtures/`` 的夹具路径（如 ``"read_state_e2e/multi_sheet.xlsx"``）。
        description: 用例描述。
        action: 不含 ``excel:`` 前缀的动作名（如 ``"get:workbookInfo"``）。
        params: snake_case 业务参数（无需 ``document_uri``）。
        validator: 成功路径验证器（可选）；双参时第二参为 :class:`WorkbookReader`。
        expect_error_code: 错误码路径——置位时**预期失败**且 error 含该码（如 ``"5001"``）。
        pre_ops: 量测动作前的预备操作 ``[(action, params), ...]``（如先 set 再 get）。
        tags: 标签。
    """

    name: str
    fixture_name: str
    description: str
    action: str
    params: dict[str, Any] = field(default_factory=dict)
    validator: Validator | None = None
    expect_error_code: str | None = None
    pre_ops: list[tuple[str, dict[str, Any]]] = field(default_factory=list)
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

            print(f"\n📝 执行: excel:{case.action} (params={case.params})...")
            start = time.time()
            ok, data, err = await excel_op(workspace, fixture.document_uri, case.action, **case.params)
            print(f"\n⏱️  执行时间: {(time.time() - start) * 1000:.1f}ms")

            # 错误码路径：预期失败 + 校验码
            if case.expect_error_code:
                passed = (not ok) and case.expect_error_code in (err or "")
                print(f"   {'✅' if passed else '❌'} 预期失败 [{case.expect_error_code}] → ok={ok} err={err}")
                _verdict(number, passed)
                return passed

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
    runner = ExcelTestRunner(
        fixtures_dir=EXCEL_FIXTURES_ROOT,
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
        return f"err {case.expect_error_code}"
    return ",".join(case.tags) if case.tags else "ok"
