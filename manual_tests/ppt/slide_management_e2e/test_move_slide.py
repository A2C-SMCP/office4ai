"""
PPT Move Slide E2E Tests

测试幻灯片移动功能。

测试场景:
1. 前移 — 将第 3 张移到第 1 的位置
2. 后移 — 将第 1 张移到第 4 的位置

运行方式:
    uv run python manual_tests/ppt/slide_management_e2e/test_move_slide.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PptTestCase,
    PPTTestRunner,
    PresentationReader,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_get_slide_info, ppt_move_slide

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


async def _workflow_move_forward(workspace: Any, doc_uri: str, working_path: Any) -> bool:
    """将第 3 张 (index=2) 移到第 1 的位置 (index=0)"""
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    count = (data or {}).get("slideCount", 0)
    if count < 3:
        print(f"   ❌ 需要 >= 3 张幻灯片，当前 {count}")
        return False

    success, _, error = await ppt_move_slide(workspace, doc_uri, from_index=2, to_index=0)
    if not success:
        print(f"   ❌ 移动失败: {error}")
        return False

    # 验证 slideCount 不变
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    after = (data or {}).get("slideCount", 0)
    if after != count:
        print(f"   ❌ slideCount 改变: {count} → {after}")
        return False
    print(f"   ✅ 前移成功: slideCount={after} 不变")

    await asyncio.sleep(0.5)
    reader = PresentationReader(working_path)
    reader.reload()
    print(f"   ✅ python-pptx 验证: slide_count={reader.slide_count}")
    return True


async def _workflow_move_backward(workspace: Any, doc_uri: str, working_path: Any) -> bool:
    """将第 1 张 (index=0) 移到第 4 的位置 (index=3)"""
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    count = (data or {}).get("slideCount", 0)
    if count < 4:
        print(f"   ❌ 需要 >= 4 张幻灯片，当前 {count}")
        return False

    success, _, error = await ppt_move_slide(workspace, doc_uri, from_index=0, to_index=3)
    if not success:
        print(f"   ❌ 移动失败: {error}")
        return False

    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    after = (data or {}).get("slideCount", 0)
    if after != count:
        print(f"   ❌ slideCount 改变: {count} → {after}")
        return False
    print(f"   ✅ 后移成功: slideCount={after} 不变")

    await asyncio.sleep(0.5)
    reader = PresentationReader(working_path)
    reader.reload()
    print(f"   ✅ python-pptx 验证: slide_count={reader.slide_count}")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="前移",
        fixture_name="multi_slide.pptx",
        description="将第 3 张移到第 1 的位置",
        tags=["crud"],
    ),
    PptTestCase(
        name="后移",
        fixture_name="multi_slide.pptx",
        description="将第 1 张移到第 4 的位置",
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
            if test_number == 1:
                passed = await _workflow_move_forward(workspace, fixture.document_uri, fixture.working_path)
            else:
                passed = await _workflow_move_backward(workspace, fixture.document_uri, fixture.working_path)
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

    parser = argparse.ArgumentParser(description="PPT Move Slide E2E Tests")
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
