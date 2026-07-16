"""
PPT Update Table Format E2E Tests

测试表格格式更新功能。

测试场景:
1. 单元格格式 — 设置单元格背景色/字体
2. 行格式 — 设置行高/背景色
3. 列格式 — 设置列宽/背景色

运行方式:
    uv run python manual_tests/ppt/update_table_e2e/test_table_format.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PPTTestRunner,
    PptTestCase,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import (
    ppt_get_current_slide_elements,
    ppt_insert_table,
    ppt_update_table_format,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def _extract_table_element_id(elements: list[dict[str, Any]]) -> str | None:
    for elem in elements:
        if "table" in elem.get("type", "").lower():
            return elem.get("id")
    return elements[-1].get("id") if elements else None


async def _setup_table(workspace: Any, doc_uri: str) -> str | None:
    success, _, error = await ppt_insert_table(workspace, doc_uri, {"rows": 3, "columns": 3})
    if not success:
        return None
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return None
    return _extract_table_element_id((data or {}).get("elements", []))


async def _workflow_cell_format(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_format(
        workspace,
        doc_uri,
        element_id,
        cell_formats=[
            {
                "rowIndex": 0,
                "columnIndex": 0,
                "backgroundColor": "#FF0000",
                # OASP 0.4.0: 字体收敛到嵌套 font（PptFont）
                "font": {"size": 16, "bold": True},
            }
        ],
    )
    if not success:
        print(f"   ❌ 单元格格式更新失败: {error}")
        return False
    print("   ✅ 单元格格式更新成功")
    return True


async def _workflow_row_format(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_format(
        workspace,
        doc_uri,
        element_id,
        row_formats=[{"rowIndex": 0, "backgroundColor": "#00FF00", "font": {"size": 14}}],
    )
    if not success:
        print(f"   ❌ 行格式更新失败: {error}")
        return False
    print("   ✅ 行格式更新成功")
    return True


async def _workflow_column_format(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_format(
        workspace,
        doc_uri,
        element_id,
        column_formats=[{"columnIndex": 0, "backgroundColor": "#0000FF", "font": {"size": 12}}],
    )
    if not success:
        print(f"   ❌ 列格式更新失败: {error}")
        return False
    print("   ✅ 列格式更新成功")
    return True


_WORKFLOW_FUNCS = [_workflow_cell_format, _workflow_row_format, _workflow_column_format]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="单元格格式", fixture_name="empty.pptx", description="设置 [0,0] 背景红色+粗体", tags=["crud"]),
    PptTestCase(name="行格式", fixture_name="empty.pptx", description="设置第一行绿色背景", tags=["crud"]),
    PptTestCase(name="列格式", fixture_name="empty.pptx", description="设置第一列蓝色背景", tags=["crud"]),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            passed = await _WORKFLOW_FUNCS[test_number - 1](workspace, fixture.document_uri)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
            return passed

    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        import traceback

        traceback.print_exc()
        return False


async def run_tests(test_indices: list[int], auto_open: bool = True, cleanup_on_success: bool = True) -> bool:
    ensure_ppt_fixtures(FIXTURES_DIR)
    runner = PPTTestRunner(fixtures_dir=FIXTURES_DIR.parent, auto_open=auto_open, cleanup_on_success=cleanup_on_success)
    results: list[bool] = []
    for idx in test_indices:
        if idx < 1 or idx > len(TEST_CASES):
            continue
        if len(test_indices) > 1 and results:
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续...")
        result = await run_single_test(runner, TEST_CASES[idx - 1], idx)
        results.append(result)
    if len(results) > 1:
        print(f"\n📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    return all(results)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="PPT Update Table Format E2E Tests")
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--always-cleanup", action="store_true")
    parser.add_argument("--list", action="store_true")
    args = parser.parse_args()

    if args.list:
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name} — {tc.description}")
        return

    test_indices = list(range(1, len(TEST_CASES) + 1)) if args.test == "all" else [int(args.test)]
    try:
        success = asyncio.run(
            run_tests(test_indices, auto_open=not args.no_auto_open, cleanup_on_success=not args.always_cleanup)
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
