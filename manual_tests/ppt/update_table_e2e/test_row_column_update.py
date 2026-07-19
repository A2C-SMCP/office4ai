"""
PPT Update Table Row/Column E2E Tests

测试表格行/列批量更新功能。

测试场景:
1. 更新行 — 更新整行数据
2. 更新列 — 更新整列数据
3. 混合更新 — 同时更新行和列

运行方式:
    uv run python manual_tests/ppt/update_table_e2e/test_row_column_update.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PptTestCase,
    PPTTestRunner,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import (
    ppt_get_current_slide_elements,
    ppt_insert_table,
    ppt_update_table_row_column,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def _extract_table_element_id(elements: list[dict[str, Any]]) -> str | None:
    for elem in elements:
        if "table" in elem.get("type", "").lower():
            return elem.get("id")
    return elements[-1].get("id") if elements else None


async def _setup_table(workspace: Any, doc_uri: str) -> str | None:
    """通用: 插入 3x3 表格并返回 elementId"""
    success, _, error = await ppt_insert_table(workspace, doc_uri, {"rows": 3, "columns": 3})
    if not success:
        print(f"   ❌ 插入表格失败: {error}")
        return None
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return None
    return _extract_table_element_id((data or {}).get("elements", []))


async def _workflow_update_row(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_row_column(
        workspace, doc_uri, element_id, rows=[{"rowIndex": 0, "values": ["H1", "H2", "H3"]}]
    )
    if not success:
        print(f"   ❌ 行更新失败: {error}")
        return False
    print("   ✅ 行更新成功 (row 0 → H1/H2/H3)")
    return True


async def _workflow_update_column(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_row_column(
        workspace, doc_uri, element_id, columns=[{"columnIndex": 0, "values": ["R1", "R2", "R3"]}]
    )
    if not success:
        print(f"   ❌ 列更新失败: {error}")
        return False
    print("   ✅ 列更新成功 (col 0 → R1/R2/R3)")
    return True


async def _workflow_mixed_update(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_table(workspace, doc_uri)
    if not element_id:
        return False

    success, _, error = await ppt_update_table_row_column(
        workspace,
        doc_uri,
        element_id,
        rows=[{"rowIndex": 0, "values": ["H1", "H2", "H3"]}],
        columns=[{"columnIndex": 2, "values": ["C1", "C2", "C3"]}],
    )
    if not success:
        print(f"   ❌ 混合更新失败: {error}")
        return False
    print("   ✅ 混合更新成功")
    return True


_WORKFLOW_FUNCS = [_workflow_update_row, _workflow_update_column, _workflow_mixed_update]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="更新行", fixture_name="empty.pptx", description="更新表格第一行为 H1/H2/H3", tags=["crud"]),
    PptTestCase(name="更新列", fixture_name="empty.pptx", description="更新表格第一列为 R1/R2/R3", tags=["crud"]),
    PptTestCase(name="混合更新", fixture_name="empty.pptx", description="同时更新第一行和最后一列", tags=["crud"]),
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

    parser = argparse.ArgumentParser(description="PPT Update Table Row/Column E2E Tests")
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
