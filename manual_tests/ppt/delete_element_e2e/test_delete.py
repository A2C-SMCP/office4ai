"""
PPT Delete Element E2E Tests

测试元素删除功能（有状态工作流: insert → get → delete → verify count）。

测试场景:
1. 单个删除 — 删除一个元素
2. 批量删除 — 删除多个元素
3. 指定幻灯片删除 — 在特定幻灯片上删除

运行方式:
    uv run python manual_tests/ppt/delete_element_e2e/test_delete.py --test all
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
    ppt_delete_element,
    ppt_get_current_slide_elements,
    ppt_insert_shape,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


async def _workflow_single_delete(workspace: Any, doc_uri: str) -> bool:
    """insert shape → get elements → delete → verify count decreased"""
    # Insert
    success, _, error = await ppt_insert_shape(
        workspace, doc_uri, "Rectangle", options={"left": 100, "top": 100, "width": 200, "height": 150}
    )
    if not success:
        print(f"   ❌ 插入形状失败: {error}")
        return False

    # Get elements and count
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    before_count = len(elements)
    element_id = elements[-1].get("id") if elements else None
    if not element_id:
        print("   ❌ 未找到元素 ID")
        return False
    print(f"   ✅ 删除前元素数: {before_count}, 目标 ID: {element_id}")

    # Delete
    success, _, error = await ppt_delete_element(workspace, doc_uri, element_id=element_id)
    if not success:
        print(f"   ❌ 删除失败: {error}")
        return False

    # Verify
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    after_count = len((data or {}).get("elements", []))
    if after_count >= before_count:
        print(f"   ❌ 删除后元素数 {after_count} 未减少")
        return False
    print(f"   ✅ 删除成功: {before_count} → {after_count}")
    return True


async def _workflow_batch_delete(workspace: Any, doc_uri: str) -> bool:
    """insert 3 shapes → get → batch delete 2 → verify"""
    for i in range(3):
        success, _, _ = await ppt_insert_shape(
            workspace, doc_uri, "Rectangle", options={"left": 50 + i * 200, "top": 100, "width": 150, "height": 100}
        )
        if not success:
            return False

    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    before_count = len(elements)
    ids_to_delete = [e.get("id") for e in elements[-2:] if e.get("id")]
    if len(ids_to_delete) < 2:
        print("   ❌ 找不到足够的元素 ID")
        return False

    success, _, error = await ppt_delete_element(workspace, doc_uri, element_ids=ids_to_delete)
    if not success:
        print(f"   ❌ 批量删除失败: {error}")
        return False

    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    after_count = len((data or {}).get("elements", []))
    print(f"   ✅ 批量删除: {before_count} → {after_count}")
    return after_count < before_count


async def _workflow_specific_slide_delete(workspace: Any, doc_uri: str) -> bool:
    """在 multi_element.pptx 第一张幻灯片上删除一个元素"""
    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    if not elements:
        print("   ❌ 当前幻灯片无元素")
        return False
    before_count = len(elements)
    element_id = elements[-1].get("id")

    success, _, error = await ppt_delete_element(workspace, doc_uri, element_id=element_id, slide_index=0)
    if not success:
        print(f"   ❌ 删除失败: {error}")
        return False

    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    after_count = len((data or {}).get("elements", []))
    print(f"   ✅ 指定幻灯片删除: {before_count} → {after_count}")
    return after_count < before_count


_WORKFLOW_FUNCS = [_workflow_single_delete, _workflow_batch_delete, _workflow_specific_slide_delete]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="单个删除", fixture_name="empty.pptx", description="插入矩形后删除", tags=["crud"]),
    PptTestCase(name="批量删除", fixture_name="empty.pptx", description="插入 3 个矩形后删除 2 个", tags=["crud"]),
    PptTestCase(
        name="指定幻灯片删除",
        fixture_name="multi_element.pptx",
        description="删除 multi_element 的一个元素",
        tags=["crud"],
    ),
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

    parser = argparse.ArgumentParser(description="PPT Delete Element E2E Tests")
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
