"""
PPT Insert Shape E2E Tests — Shape Types

测试不同形状类型的插入。

测试场景:
1. 批量基本形状 — Rectangle, Circle, Oval, Triangle
2. 箭头线条 — Arrow, Line
3. 星形文本框 — Star, RoundedRectangle, TextBox

运行方式:
    uv run python manual_tests/ppt/insert_shape_e2e/test_shape_types.py --test all
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
from manual_tests.ppt.test_helpers import ppt_insert_shape

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_batch_shapes(data: dict[str, Any]) -> bool:
    """验证批量基本形状插入（最后一个的结果）"""
    print(f"   ✅ 最后一个形状插入返回: {data}")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="批量基本形状",
        fixture_name="empty.pptx",
        description="连续插入 Rectangle, Circle, Oval, Triangle",
        validator=validate_batch_shapes,
        tags=["basic"],
    ),
    PptTestCase(
        name="箭头线条",
        fixture_name="empty.pptx",
        description="插入 Arrow 和 Line",
        validator=validate_batch_shapes,
        tags=["basic"],
    ),
    PptTestCase(
        name="星形文本框",
        fixture_name="empty.pptx",
        description="插入 Star, RoundedRectangle, TextBox",
        validator=validate_batch_shapes,
        tags=["advanced"],
    ),
]

_SHAPE_GROUPS: list[list[tuple[str, dict[str, Any]]]] = [
    [
        ("Rectangle", {"left": 50, "top": 50, "width": 150, "height": 100}),
        ("Circle", {"left": 250, "top": 50, "width": 100, "height": 100}),
        ("Oval", {"left": 400, "top": 50, "width": 150, "height": 100}),
        ("Triangle", {"left": 600, "top": 50, "width": 120, "height": 100}),
    ],
    [
        ("Arrow", {"left": 100, "top": 100, "width": 200, "height": 50}),
        ("Line", {"left": 100, "top": 250, "width": 300, "height": 5}),
    ],
    [
        ("Star", {"left": 50, "top": 50, "width": 150, "height": 150}),
        ("RoundedRectangle", {"left": 250, "top": 50, "width": 150, "height": 100}),
        ("TextBox", {"left": 450, "top": 50, "width": 200, "height": 80, "text": "Hello"}),
    ],
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    shapes = _SHAPE_GROUPS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            last_data: dict[str, Any] = {}

            for shape_type, options in shapes:
                print(f"\n   📝 插入 {shape_type}...")
                success, data, error = await ppt_insert_shape(
                    workspace, fixture.document_uri, shape_type, options=options
                )
                if not success:
                    print(f"   ❌ 插入 {shape_type} 失败: {error}")
                    return False
                print(f"   ✅ {shape_type} 插入成功")
                last_data = data or {}

            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")

            print("\n📊 验证结果:")
            passed = True
            if test_case.validator and not test_case.validator(last_data):
                passed = False

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

    parser = argparse.ArgumentParser(description="PPT Insert Shape Types E2E Tests")
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
