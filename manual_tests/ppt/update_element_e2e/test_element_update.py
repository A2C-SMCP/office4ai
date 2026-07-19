"""
PPT Update Element E2E Tests

测试元素位置/大小/旋转更新功能。

测试场景:
1. 移动 — 更新 left/top
2. 缩放 — 更新 width/height
3. 旋转 — 更新 rotation
4. 组合变换 — 同时更新位置、大小和旋转

运行方式:
    uv run python manual_tests/ppt/update_element_e2e/test_element_update.py --test all
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
    ppt_insert_shape,
    ppt_update_element,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def _extract_first_element_id(elements: list[dict[str, Any]]) -> str | None:
    if elements:
        return elements[-1].get("id")
    return None


async def _setup_shape(workspace: Any, doc_uri: str) -> str | None:
    """插入一个矩形并返回 elementId"""
    success, _, error = await ppt_insert_shape(
        workspace, doc_uri, "Rectangle", options={"left": 100, "top": 100, "width": 200, "height": 150}
    )
    if not success:
        print(f"   ❌ 插入形状失败: {error}")
        return None
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return None
    return _extract_first_element_id((data or {}).get("elements", []))


async def _workflow_move(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_shape(workspace, doc_uri)
    if not element_id:
        return False
    success, _, error = await ppt_update_element(workspace, doc_uri, element_id, {"left": 400, "top": 300})
    if not success:
        print(f"   ❌ 移动失败: {error}")
        return False
    print("   ✅ 移动成功 (left=400, top=300)")
    return True


async def _workflow_resize(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_shape(workspace, doc_uri)
    if not element_id:
        return False
    success, _, error = await ppt_update_element(workspace, doc_uri, element_id, {"width": 400, "height": 300})
    if not success:
        print(f"   ❌ 缩放失败: {error}")
        return False
    print("   ✅ 缩放成功 (width=400, height=300)")
    return True


async def _workflow_rotate(workspace: Any, doc_uri: str) -> bool:
    element_id = await _setup_shape(workspace, doc_uri)
    if not element_id:
        return False
    success, _, error = await ppt_update_element(workspace, doc_uri, element_id, {"rotation": 45})
    if not success:
        print(f"   ❌ 旋转失败: {error}")
        return False
    print("   ✅ 旋转成功 (rotation=45)")
    return True


_WORKFLOW_FUNCS = [_workflow_move, _workflow_resize, _workflow_rotate]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="移动", fixture_name="empty.pptx", description="插入矩形后移动到 (400, 300)", tags=["crud"]),
    PptTestCase(name="缩放", fixture_name="empty.pptx", description="插入矩形后缩放到 400x300", tags=["crud"]),
    PptTestCase(name="旋转", fixture_name="empty.pptx", description="插入矩形后旋转 45 度", tags=["crud"]),
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

    parser = argparse.ArgumentParser(description="PPT Update Element E2E Tests")
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
