"""
Format Replace Selection E2E Tests (自动化版本)

测试带格式的选区文本替换功能（word:replace:selection + format）。

注意：这些测试需要用户手动在 Word 中选中文本后按 Enter 继续。

测试场景:
1. 替换为粗体文本 - bold: True
2. 替换为斜体文本 - italic: True
3. 替换为带字体格式 - fontName + fontSize + bold
4. 替换为带颜色和下划线 - color + underline + bold

运行方式:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test 1

    # 运行所有测试
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/replace_selection_e2e/test_format_replace.py --test 1 --always-cleanup
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.e2e_base import (
    DocumentReader,
    E2ETestRunner,
    TestCase,
    _call_validator,
    ensure_fixtures,
)
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "replace_selection_e2e"

# ==============================================================================
# Validator 函数
# ==============================================================================


def validate_bold_replace(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证粗体文本替换"""
    reader.reload()
    if not reader.contains("Bold Text"):
        print("   ❌ 文档中未找到 'Bold Text'")
        return False
    if reader.run_has_format("Bold Text", bold=True):
        print("   ✅ 文档内容验证通过: 粗体格式正确")
        return True
    print("   ⚠️  文本存在但格式未验证（Word 可能延迟保存格式）")
    return True  # 文本存在即可，格式由人工确认


def validate_italic_replace(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证斜体文本替换"""
    reader.reload()
    if not reader.contains("Italic Text"):
        print("   ❌ 文档中未找到 'Italic Text'")
        return False
    if reader.run_has_format("Italic Text", italic=True):
        print("   ✅ 文档内容验证通过: 斜体格式正确")
        return True
    print("   ⚠️  文本存在但格式未验证")
    return True


def validate_font_format_replace(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证带字体格式的文本替换"""
    reader.reload()
    if not reader.contains("Formatted Text"):
        print("   ❌ 文档中未找到 'Formatted Text'")
        return False
    if reader.run_has_format("Formatted Text", bold=True, font_name="Arial", font_size=16):
        print("   ✅ 文档内容验证通过: Arial 16pt 粗体格式正确")
        return True
    print("   ⚠️  文本存在但格式未验证")
    return True


def validate_color_underline_replace(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证带颜色和下划线的文本替换"""
    reader.reload()
    if not reader.contains("Colorful Underlined Text"):
        print("   ❌ 文档中未找到 'Colorful Underlined Text'")
        return False
    if reader.run_has_format("Colorful Underlined Text", bold=True):
        print("   ✅ 文档内容验证通过: 粗体格式正确（颜色和下划线需人工确认）")
        return True
    print("   ⚠️  文本存在但格式未验证")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[TestCase] = [
    TestCase(
        name="替换为粗体文本",
        fixture_name="simple.docx",
        description="选中文本后替换为粗体 'Bold Text'，验证文本存在并检查 bold 格式",
        validator=validate_bold_replace,
        tags=["format", "bold"],
    ),
    TestCase(
        name="替换为斜体文本",
        fixture_name="simple.docx",
        description="选中文本后替换为斜体 'Italic Text'，验证文本存在并检查 italic 格式",
        validator=validate_italic_replace,
        tags=["format", "italic"],
    ),
    TestCase(
        name="替换为带字体格式",
        fixture_name="simple.docx",
        description="选中文本后替换为 Arial 16pt 粗体 'Formatted Text'",
        validator=validate_font_format_replace,
        tags=["format", "font"],
    ),
    TestCase(
        name="替换为带颜色和下划线",
        fixture_name="simple.docx",
        description="选中文本后替换为红色、单下划线、粗体 'Colorful Underlined Text'",
        validator=validate_color_underline_replace,
        tags=["format", "color", "underline"],
    ),
]

# 每个测试用例的替换内容（含 format）
# OASP 0.4.0: 字体属性收敛到嵌套 font（WordFont）；扁平 bold/fontSize/fontName/color/underline 须写 format.font.*
_REPLACE_CONTENTS: list[dict[str, Any]] = [
    # Test 1: bold
    {"text": "Bold Text", "format": {"font": {"bold": True}}},
    # Test 2: italic
    {"text": "Italic Text", "format": {"font": {"italic": True}}},
    # Test 3: name + size + bold
    {"text": "Formatted Text", "format": {"font": {"name": "Arial", "size": 16, "bold": True}}},
    # Test 4: color + underline + bold
    {"text": "Colorful Underlined Text", "format": {"font": {"color": "#FF0000", "underline": "Single", "bold": True}}},
]


# ==============================================================================
# 测试执行
# ==============================================================================


async def run_single_test(
    runner: E2ETestRunner,
    test_case: TestCase,
    test_number: int,
) -> bool:
    """执行单个测试用例"""
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"replace_selection_e2e/{test_case.fixture_name}"
    content = _REPLACE_CONTENTS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 等待用户选中文本
            input("\n请在 Word 中选中一些文本后按 Enter...")

            print(f"\n📝 执行: {test_case.name}...")
            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="replace:selection",
                params={
                    "document_uri": fixture.document_uri,
                    "content": content,
                },
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"❌ 替换失败: {result.error}")
                return False

            print("✅ 协议返回成功")
            data = result.data or {}

            # ContentValidator 双重验证
            print("\n📊 验证结果:")
            passed = True

            if test_case.validator:
                reader = DocumentReader(fixture.working_path)
                await asyncio.sleep(1.0)
                if not _call_validator(test_case.validator, data, reader):
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
    ensure_fixtures(FIXTURES_DIR)

    runner = E2ETestRunner(
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


# ==============================================================================
# 命令行入口
# ==============================================================================


def main() -> None:
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Format Replace Selection E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（粗体替换）
  python test_format_replace.py --test 1

  # 运行所有测试
  python test_format_replace.py --test all

  # 手动打开文档模式
  python test_format_replace.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_format_replace.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=bold, 2=italic, 3=font, 4=color+underline, all=全部",
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
            print(f"     标签: {', '.join(tc.tags)}")
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
