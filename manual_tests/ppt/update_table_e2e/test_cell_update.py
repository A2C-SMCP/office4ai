"""
PPT Update Table Cell E2E Tests

测试表格单元格更新功能。

测试场景:
1. 单格更新 — 更新单个单元格文本
2. 多格更新 — 批量更新多个单元格
3. 边角格更新 — 更新表格四角的单元格

运行方式:
    uv run python manual_tests/ppt/update_table_e2e/test_cell_update.py --test all
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
    ppt_update_table_cell,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def _extract_table_element_id(elements: list[dict[str, Any]]) -> str | None:
    for elem in elements:
        elem_type = elem.get("type", "").lower()
        if "table" in elem_type:
            return elem.get("id")
    if elements:
        return elements[-1].get("id")
    return None


async def _workflow_single_cell(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert table → get → update single cell"""
    print("\n   📝 Step 1: 插入 3x3 表格...")
    success, _, error = await ppt_insert_table(workspace, doc_uri, {"rows": 3, "columns": 3})
    if not success:
        print(f"   ❌ 插入表格失败: {error}")
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_table_element_id(elements)
    if not element_id:
        print("   ❌ 未找到表格元素 ID")
        return False

    print(f"   📝 Step 2: 更新 [0,0] → '已更新'...")
    success, _, error = await ppt_update_table_cell(
        workspace, doc_uri, element_id, [{"rowIndex": 0, "columnIndex": 0, "text": "已更新"}]
    )
    if not success:
        print(f"   ❌ 更新失败: {error}")
        return False
    print("   ✅ 单格更新成功")
    return True


async def _workflow_multi_cell(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update multiple cells"""
    success, _, error = await ppt_insert_table(workspace, doc_uri, {"rows": 3, "columns": 3})
    if not success:
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    element_id = _extract_table_element_id((data or {}).get("elements", []))
    if not element_id:
        return False

    cells = [
        {"rowIndex": 0, "columnIndex": 0, "text": "A1"},
        {"rowIndex": 1, "columnIndex": 1, "text": "B2"},
        {"rowIndex": 2, "columnIndex": 2, "text": "C3"},
    ]
    success, _, error = await ppt_update_table_cell(workspace, doc_uri, element_id, cells)
    if not success:
        print(f"   ❌ 多格更新失败: {error}")
        return False
    print("   ✅ 多格更新成功 (对角线 A1/B2/C3)")
    return True


async def _workflow_corner_cells(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update corner cells"""
    success, _, error = await ppt_insert_table(workspace, doc_uri, {"rows": 3, "columns": 3})
    if not success:
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    element_id = _extract_table_element_id((data or {}).get("elements", []))
    if not element_id:
        return False

    cells = [
        {"rowIndex": 0, "columnIndex": 0, "text": "左上"},
        {"rowIndex": 0, "columnIndex": 2, "text": "右上"},
        {"rowIndex": 2, "columnIndex": 0, "text": "左下"},
        {"rowIndex": 2, "columnIndex": 2, "text": "右下"},
    ]
    success, _, error = await ppt_update_table_cell(workspace, doc_uri, element_id, cells)
    if not success:
        print(f"   ❌ 边角格更新失败: {error}")
        return False
    print("   ✅ 边角格更新成功")
    return True


_WORKFLOW_FUNCS = [_workflow_single_cell, _workflow_multi_cell, _workflow_corner_cells]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="单格更新", fixture_name="empty.pptx", description="更新表格 [0,0] 单元格", tags=["crud"]),
    PptTestCase(name="多格更新", fixture_name="empty.pptx", description="批量更新对角线单元格", tags=["crud"]),
    PptTestCase(name="边角格更新", fixture_name="empty.pptx", description="更新表格四角单元格", tags=["crud"]),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            passed = await _WORKFLOW_FUNCS[test_number - 1](workspace, fixture.document_uri)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")
            print("\n" + "=" * 70)
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
            print("=" * 70)
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

    parser = argparse.ArgumentParser(description="PPT Update Table Cell E2E Tests")
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
