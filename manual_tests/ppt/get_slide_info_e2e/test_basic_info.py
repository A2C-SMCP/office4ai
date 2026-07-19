"""
PPT Get Slide Info E2E Tests

测试获取幻灯片基本信息功能。

测试场景:
1. 单幻灯片信息 — 空白 PPT 的基本信息
2. 多幻灯片信息 — 3 页 PPT 的幻灯片数量
3. 幻灯片尺寸 — 验证默认 PPT 的宽高
4. 指定幻灯片详情 — 通过 slideIndex 获取特定幻灯片信息

运行方式:
    uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --test 1
    uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --test all
    uv run python manual_tests/ppt/get_slide_info_e2e/test_basic_info.py --list
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
    _call_ppt_validator,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_get_slide_info

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"

# ==============================================================================
# Validators
# ==============================================================================


def validate_single_slide(data: dict[str, Any]) -> bool:
    """验证单幻灯片 PPT 信息"""
    slide_count = data.get("slideCount", 0)
    if slide_count < 1:
        print(f"   ❌ slideCount 应 >= 1，实际: {slide_count}")
        return False
    print(f"   ✅ slideCount={slide_count}")
    return True


def validate_multi_slide(data: dict[str, Any], reader: PresentationReader) -> bool:
    """验证多幻灯片 PPT 信息（双重验证）"""
    slide_count = data.get("slideCount", 0)
    if slide_count < 3:
        print(f"   ❌ slideCount 应 >= 3，实际: {slide_count}")
        return False
    print(f"   ✅ 协议返回 slideCount={slide_count}")

    reader.reload()
    actual_count = reader.slide_count
    if actual_count < 3:
        print(f"   ❌ python-pptx 验证 slide_count={actual_count}，应 >= 3")
        return False
    print(f"   ✅ python-pptx 验证 slide_count={actual_count}")
    return True


def validate_slide_dimensions(data: dict[str, Any]) -> bool:
    """验证幻灯片尺寸（默认 10x7.5 inches = 720x540 points）"""
    dimensions = data.get("dimensions", {})
    width = dimensions.get("width", 0)
    height = dimensions.get("height", 0)
    if width <= 0 or height <= 0:
        print(f"   ❌ 尺寸无效: width={width}, height={height}")
        return False
    aspect = dimensions.get("aspectRatio", "unknown")
    print(f"   ✅ 幻灯片尺寸: width={width}, height={height}, aspectRatio={aspect}")
    return True


def validate_specific_slide(data: dict[str, Any]) -> bool:
    """验证指定幻灯片详情"""
    # 至少应返回 slideCount 和 currentSlideIndex
    if "slideCount" not in data:
        print("   ❌ 缺少 slideCount 字段")
        return False
    print(f"   ✅ slideCount={data.get('slideCount')}")
    return True


# ==============================================================================
# 测试用例
# ==============================================================================

TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="单幻灯片信息",
        fixture_name="empty.pptx",
        description="获取空白 PPT 的基本信息，验证 slideCount >= 1",
        validator=validate_single_slide,
        tags=["basic"],
    ),
    PptTestCase(
        name="多幻灯片信息",
        fixture_name="simple.pptx",
        description="获取 3 页 PPT 的信息，双重验证幻灯片数量",
        validator=validate_multi_slide,
        tags=["basic"],
    ),
    PptTestCase(
        name="幻灯片尺寸",
        fixture_name="simple.pptx",
        description="获取 PPT 信息，验证幻灯片宽高有效",
        validator=validate_slide_dimensions,
        tags=["basic"],
    ),
    PptTestCase(
        name="指定幻灯片详情",
        fixture_name="simple.pptx",
        description="通过 slideIndex=1 获取第二张幻灯片的详情",
        validator=validate_specific_slide,
        tags=["basic"],
    ),
]

# 每个测试对应的 slide_index 参数（None 表示不传）
_SLIDE_INDICES: list[int | None] = [None, None, None, 1]


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
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    slide_index = _SLIDE_INDICES[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            print(f"\n📝 执行: 获取幻灯片信息 (slideIndex={slide_index})...")
            start_time = time.time()

            success, data, error = await ppt_get_slide_info(
                workspace, fixture.document_uri, slide_index=slide_index
            )
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 获取失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}
            print(f"   数据: {data}")

            # 验证
            print("\n📊 验证结果:")
            passed = True

            if test_case.expected:
                verified, messages = runner.verify_slide_info(data, test_case.expected)
                for msg in messages:
                    print(f"   {msg}")
                if not verified:
                    passed = False

            if test_case.validator:
                await asyncio.sleep(0.5)
                reader = PresentationReader(fixture.working_path)
                if not _call_ppt_validator(test_case.validator, data, reader):
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

    parser = argparse.ArgumentParser(
        description="PPT Get Slide Info E2E Tests",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试编号或 all",
    )
    parser.add_argument("--no-auto-open", action="store_true", help="不自动打开文档")
    parser.add_argument("--always-cleanup", action="store_true", help="无论成功失败都清理")
    parser.add_argument("--list", action="store_true", help="列出所有测试用例")

    args = parser.parse_args()

    if args.list:
        print("\n📋 可用测试用例:\n")
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name}")
            print(f"     夹具: {tc.fixture_name}")
            print(f"     描述: {tc.description}")
            print()
        return

    if args.test == "all":
        test_indices = list(range(1, len(TEST_CASES) + 1))
    else:
        test_indices = [int(args.test)]

    try:
        success = asyncio.run(
            run_tests(
                test_indices=test_indices,
                auto_open=not args.no_auto_open,
                cleanup_on_success=not args.always_cleanup,
            )
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
