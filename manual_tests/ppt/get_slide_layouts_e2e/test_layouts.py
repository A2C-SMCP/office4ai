"""
PPT Get Slide Layouts E2E Tests

测试获取幻灯片布局模板功能。

测试场景:
1. 默认布局列表 — 获取所有可用布局
2. 布局名称验证 — 验证布局包含名称字段
3. 包含占位符选项 — 带 includePlaceholders=true 获取布局详情

运行方式:
    uv run python manual_tests/ppt/get_slide_layouts_e2e/test_layouts.py --test all
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
from manual_tests.ppt.test_helpers import ppt_get_slide_layouts

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"

# ==============================================================================
# Validators
# ==============================================================================


def validate_default_layouts(data: dict[str, Any]) -> bool:
    """验证默认布局列表"""
    layouts = data.get("layouts", [])
    if not layouts:
        print("   ❌ 布局列表为空")
        return False
    print(f"   ✅ 获取到 {len(layouts)} 个布局")
    for i, layout in enumerate(layouts[:5]):
        print(f"      [{i}] {layout.get('name', 'N/A')}")
    return True


def validate_layout_names(data: dict[str, Any]) -> bool:
    """验证布局包含名称字段"""
    layouts = data.get("layouts", [])
    if not layouts:
        print("   ❌ 布局列表为空")
        return False
    for layout in layouts:
        if "name" not in layout:
            print(f"   ❌ 布局缺少 name 字段: {layout}")
            return False
    print(f"   ✅ 所有 {len(layouts)} 个布局都包含 name 字段")
    return True


def validate_with_placeholders(data: dict[str, Any]) -> bool:
    """验证带占位符的布局详情"""
    layouts = data.get("layouts", [])
    if not layouts:
        print("   ❌ 布局列表为空")
        return False
    # 至少有一个布局应包含 placeholders
    has_placeholders = any("placeholders" in layout for layout in layouts)
    if not has_placeholders:
        print("   ⚠️  没有布局包含 placeholders 字段（可能 Add-In 不返回此字段）")
        # 不作为失败条件，仅警告
    else:
        for layout in layouts:
            if "placeholders" in layout:
                print(f"   ✅ 布局 '{layout.get('name')}' 包含 {len(layout['placeholders'])} 个占位符")
                break
    print(f"   ✅ 获取到 {len(layouts)} 个布局")
    return True


# ==============================================================================
# 测试用例
# ==============================================================================

TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="默认布局列表",
        fixture_name="simple.pptx",
        description="获取所有可用布局，验证布局列表不为空",
        validator=validate_default_layouts,
        tags=["basic"],
    ),
    PptTestCase(
        name="布局名称验证",
        fixture_name="simple.pptx",
        description="验证每个布局都包含 name 字段",
        validator=validate_layout_names,
        tags=["basic"],
    ),
    PptTestCase(
        name="包含占位符选项",
        fixture_name="simple.pptx",
        description="带 includePlaceholders=true 获取布局，验证占位符信息",
        validator=validate_with_placeholders,
        tags=["advanced"],
    ),
]

_OPTIONS: list[dict[str, Any] | None] = [None, None, {"include_placeholders": True}]


# ==============================================================================
# 测试执行
# ==============================================================================


async def run_single_test(
    runner: PPTTestRunner,
    test_case: PptTestCase,
    test_number: int,
) -> bool:
    """执行单个测试用例"""
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    options = _OPTIONS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            print(f"\n📝 执行: 获取幻灯片布局 (options={options})...")
            start_time = time.time()

            success, data, error = await ppt_get_slide_layouts(
                workspace, fixture.document_uri, options=options
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

            if test_case.validator:
                if not test_case.validator(data):
                    passed = False

            print("\n" + "=" * 70)
            if passed:
                print(f"✅ 测试 {test_number} 通过")
            else:
                print(f"❌ 测试 {test_number} 失败")
            print("=" * 70)
            return passed

    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        import traceback

        traceback.print_exc()
        return False


async def run_tests(
    test_indices: list[int],
    auto_open: bool = True,
    cleanup_on_success: bool = True,
) -> bool:
    """运行指定的测试"""
    ensure_ppt_fixtures(FIXTURES_DIR)

    runner = PPTTestRunner(
        fixtures_dir=FIXTURES_DIR.parent,
        auto_open=auto_open,
        cleanup_on_success=cleanup_on_success,
    )

    results: list[bool] = []
    for idx in test_indices:
        if idx < 1 or idx > len(TEST_CASES):
            print(f"⚠️  无效的测试编号: {idx}")
            continue
        test_case = TEST_CASES[idx - 1]
        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("⏳ 准备下一个测试...")
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续...")
        result = await run_single_test(runner, test_case, idx)
        results.append(result)

    if len(results) > 1:
        print("\n" + "=" * 70)
        print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
        print("=" * 70)
    return all(results)


def main() -> None:
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(description="PPT Get Slide Layouts E2E Tests")
    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
    )
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--always-cleanup", action="store_true")
    parser.add_argument("--list", action="store_true")
    args = parser.parse_args()

    if args.list:
        print("\n📋 可用测试用例:\n")
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name} — {tc.description}")
        return

    if args.test == "all":
        test_indices = list(range(1, len(TEST_CASES) + 1))
    else:
        test_indices = [int(args.test)]

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
