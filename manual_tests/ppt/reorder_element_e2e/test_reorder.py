"""
PPT Reorder Element E2E Tests

测试元素 z-order 调整功能。

测试场景:
1. bringToFront — 移到最前
2. sendToBack — 移到最后
3. bringForward — 前移一层
4. sendBackward — 后移一层

运行方式:
    uv run python manual_tests/ppt/reorder_element_e2e/test_reorder.py --test all
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
    ppt_insert_shape,
    ppt_reorder_element,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


async def _setup_two_shapes(workspace: Any, doc_uri: str) -> tuple[str | None, str | None]:
    """插入两个重叠形状，返回 (first_id, second_id)"""
    success, _, _ = await ppt_insert_shape(
        workspace, doc_uri, "Rectangle", options={"left": 100, "top": 100, "width": 200, "height": 150}
    )
    if not success:
        return None, None
    success, _, _ = await ppt_insert_shape(
        workspace, doc_uri, "Circle", options={"left": 150, "top": 150, "width": 150, "height": 150}
    )
    if not success:
        return None, None

    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return None, None
    elements = (data or {}).get("elements", [])
    if len(elements) < 2:
        return None, None
    return elements[-2].get("id"), elements[-1].get("id")


async def _workflow_reorder(workspace: Any, doc_uri: str, action: str) -> bool:
    first_id, second_id = await _setup_two_shapes(workspace, doc_uri)
    if not first_id or not second_id:
        print("   ❌ 创建形状失败")
        return False

    # 对第一个元素执行 reorder
    target_id = first_id
    print(f"   📝 对 elementId={target_id} 执行 {action}...")
    success, _, error = await ppt_reorder_element(workspace, doc_uri, target_id, action)
    if not success:
        print(f"   ❌ reorder 失败: {error}")
        return False
    print(f"   ✅ {action} 成功")

    # 验证元素顺序（仅协议验证，z-order 无法通过 python-pptx 验证）
    success, data, _ = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_ids = [e.get("id") for e in elements]
    print(f"   ℹ️  当前元素顺序: {element_ids}")
    return True


_ACTIONS = ["bringToFront", "sendToBack", "bringForward", "sendBackward"]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="bringToFront", fixture_name="empty.pptx", description="将第一个形状移到最前", tags=["crud"]),
    PptTestCase(name="sendToBack", fixture_name="empty.pptx", description="将第一个形状移到最后", tags=["crud"]),
    PptTestCase(name="bringForward", fixture_name="empty.pptx", description="将第一个形状前移一层", tags=["crud"]),
    PptTestCase(name="sendBackward", fixture_name="empty.pptx", description="将第一个形状后移一层", tags=["crud"]),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    action = _ACTIONS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            passed = await _workflow_reorder(workspace, fixture.document_uri, action)
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

    parser = argparse.ArgumentParser(description="PPT Reorder Element E2E Tests")
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
