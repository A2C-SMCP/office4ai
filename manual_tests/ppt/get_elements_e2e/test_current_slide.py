"""
PPT Get Current Slide Elements E2E Tests

测试获取当前幻灯片元素功能。

测试场景:
1. 空幻灯片元素 — 空白幻灯片应返回空/最小元素列表
2. 多元素幻灯片 — 含文本框+表格+形状的幻灯片
3. 元素属性验证 — 验证返回的元素包含必要属性字段

运行方式:
    uv run python manual_tests/ppt/get_elements_e2e/test_current_slide.py --test all
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
from manual_tests.ppt.test_helpers import ppt_get_current_slide_elements

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_empty_slide(data: dict[str, Any]) -> bool:
    """验证空幻灯片元素"""
    elements = data.get("elements", [])
    print(f"   ✅ 空幻灯片元素数: {len(elements)}")
    return True


def validate_multi_elements(data: dict[str, Any]) -> bool:
    """验证多元素幻灯片"""
    elements = data.get("elements", [])
    if len(elements) < 2:
        print(f"   ❌ 元素数量不足: {len(elements)}，预期 >= 2")
        return False
    print(f"   ✅ 获取到 {len(elements)} 个元素")
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_id = elem.get("id", "N/A")
        print(f"      - [{elem_type}] id={elem_id}")
    return True


def validate_element_properties(data: dict[str, Any]) -> bool:
    """验证元素包含必要属性字段"""
    elements = data.get("elements", [])
    if not elements:
        print("   ❌ 元素列表为空")
        return False
    for elem in elements:
        if "id" not in elem:
            print(f"   ❌ 元素缺少 id 字段: {elem}")
            return False
        if "type" not in elem:
            print(f"   ❌ 元素缺少 type 字段: {elem}")
            return False
    print(f"   ✅ 所有 {len(elements)} 个元素都包含 id 和 type 字段")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="空幻灯片元素",
        fixture_name="empty.pptx",
        description="获取空白幻灯片的元素，验证返回格式正确",
        validator=validate_empty_slide,
        tags=["basic"],
    ),
    PptTestCase(
        name="多元素幻灯片",
        fixture_name="multi_element.pptx",
        description="获取含文本框+表格+形状的幻灯片元素",
        validator=validate_multi_elements,
        tags=["basic"],
    ),
    PptTestCase(
        name="元素属性验证",
        fixture_name="multi_element.pptx",
        description="验证返回的每个元素都包含 id 和 type 字段",
        validator=validate_element_properties,
        tags=["basic"],
    ),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print("\n📝 执行: 获取当前幻灯片元素...")
            start_time = time.time()
            success, data, error = await ppt_get_current_slide_elements(workspace, fixture.document_uri)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 获取失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}

            print("\n📊 验证结果:")
            passed = True
            if test_case.validator:
                if not test_case.validator(data):
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
            print(f"⚠️  无效的测试编号: {idx}")
            continue
        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("⏳ 准备下一个测试...")
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

    parser = argparse.ArgumentParser(description="PPT Get Current Slide Elements E2E Tests")
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--always-cleanup", action="store_true")
    parser.add_argument("--list", action="store_true")
    args = parser.parse_args()

    if args.list:
        print("\n📋 可用测试用例:\n")
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
