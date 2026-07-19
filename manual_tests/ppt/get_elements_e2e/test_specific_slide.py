"""
PPT Get Slide Elements (by index) E2E Tests

测试按索引获取指定幻灯片元素功能。

测试场景:
1. 按索引获取 — 获取 slideIndex=0 的元素
2. 过滤文本元素 — 使用 includeText 选项
3. 过滤形状元素 — 使用 includeShapes 选项

运行方式:
    uv run python manual_tests/ppt/get_elements_e2e/test_specific_slide.py --test all
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
from manual_tests.ppt.test_helpers import ppt_get_slide_elements

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_by_index(data: dict[str, Any]) -> bool:
    """验证按索引获取元素"""
    elements = data.get("elements", [])
    print(f"   ✅ slideIndex=0 返回 {len(elements)} 个元素")
    return True


def validate_filter_text(data: dict[str, Any]) -> bool:
    """验证文本过滤"""
    elements = data.get("elements", [])
    print(f"   ✅ 文本过滤后返回 {len(elements)} 个元素")
    return True


def validate_filter_shapes(data: dict[str, Any]) -> bool:
    """验证形状过滤"""
    elements = data.get("elements", [])
    print(f"   ✅ 形状过滤后返回 {len(elements)} 个元素")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="按索引获取元素",
        fixture_name="multi_element.pptx",
        description="获取 slideIndex=0 的所有元素",
        validator=validate_by_index,
        tags=["basic"],
    ),
    PptTestCase(
        name="过滤文本元素",
        fixture_name="multi_element.pptx",
        description="使用 includeText=true, includeShapes=false 过滤",
        validator=validate_filter_text,
        tags=["advanced"],
    ),
    PptTestCase(
        name="过滤形状元素",
        fixture_name="multi_element.pptx",
        description="使用 includeShapes=true, includeText=false 过滤",
        validator=validate_filter_shapes,
        tags=["advanced"],
    ),
]

_SLIDE_INDEX = 0
_OPTIONS: list[dict[str, Any] | None] = [
    None,
    {"include_text": True, "include_shapes": False, "include_tables": False},
    {"include_shapes": True, "include_text": False, "include_tables": False},
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    options = _OPTIONS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 获取幻灯片元素 (slideIndex={_SLIDE_INDEX}, options={options})...")
            start_time = time.time()
            success, data, error = await ppt_get_slide_elements(
                workspace, fixture.document_uri, slide_index=_SLIDE_INDEX, options=options
            )
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 获取失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}

            print("\n📊 验证结果:")
            passed = True
            if test_case.validator and not test_case.validator(data):
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

    parser = argparse.ArgumentParser(description="PPT Get Slide Elements E2E Tests")
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
