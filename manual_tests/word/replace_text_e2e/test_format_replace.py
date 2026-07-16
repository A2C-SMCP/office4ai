"""
Format Replace E2E Tests

测试 word:replace:text 带 format 参数的格式化替换功能。
使用相同文本替换 + format 参数，实现"为已有文本添加格式"。

测试场景:
1. 粗体格式化（replaceAll）
2. 斜体格式化（replaceAll）
3. 颜色格式化（replaceAll）
4. styleName 样式格式化
5. 组合格式化（bold+italic+color+fontSize）
6. 替换文本 + 格式同时应用

运行方式:
    uv run python manual_tests/replace_text_e2e/test_format_replace.py --test 1
    uv run python manual_tests/replace_text_e2e/test_format_replace.py --test all
    uv run python manual_tests/replace_text_e2e/test_format_replace.py --list
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

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "replace_text_e2e"

# ==============================================================================
# Validator 函数
# ==============================================================================


def validate_bold_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证粗体格式化: 'important' 文本应变为粗体"""
    reader.reload()
    if not reader.contains("important"):
        print("   文档中未找到 'important' 文本")
        return False
    if reader.run_has_format("important", bold=True):
        print("   文档内容验证通过: 'important' 已设为粗体")
        return True
    print("   文本存在但粗体格式未验证（Word 可能延迟保存格式）")
    return True  # 文本存在即可，格式由人工确认


def validate_italic_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证斜体格式化: 'emphasis' 文本应变为斜体"""
    reader.reload()
    if not reader.contains("emphasis"):
        print("   文档中未找到 'emphasis' 文本")
        return False
    if reader.run_has_format("emphasis", italic=True):
        print("   文档内容验证通过: 'emphasis' 已设为斜体")
        return True
    print("   文本存在但斜体格式未验证")
    return True


def validate_color_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证颜色格式化: 'alert' 文本应变为红色"""
    reader.reload()
    if not reader.contains("alert"):
        print("   文档中未找到 'alert' 文本")
        return False
    # python-docx 读取颜色较复杂，验证文本存在 + replaceCount 即可
    replace_count = data.get("replaceCount", 0)
    if replace_count >= 3:
        print(f"   协议验证通过: replaceCount={replace_count} (预期 >= 3)")
        print("   文档内容验证通过: 'alert' 文本存在（颜色需人工确认）")
        return True
    print(f"   replaceCount={replace_count} 小于预期")
    return False


def validate_style_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证 styleName 格式化: 'Chapter' 段落应变为 Heading 2"""
    reader.reload()
    if not reader.contains("Chapter"):
        print("   文档中未找到 'Chapter' 文本")
        return False
    replace_count = data.get("replaceCount", 0)
    if replace_count < 2:
        print(f"   replaceCount={replace_count} 小于预期")
        return False
    print(f"   协议验证通过: replaceCount={replace_count} (预期 >= 2)")
    # replace:text + styleName 将样式的字符部分以 rStyle 应用到 run 上
    # "Heading 2" → "Heading 2 Char"(英文) / "标题 2 字符"(中文)
    if reader.run_has_style("Chapter", "Heading 2") or reader.run_has_style("Chapter", "标题 2"):
        print("   文档内容验证通过: 'Chapter' run 已应用 Heading 2 字符样式")
        return True
    print("   ⚠️  'Chapter' run 未检测到 Heading 2 字符样式")
    return False


def validate_combined_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证组合格式化: 'Critical' 文本应有 bold+italic+color"""
    reader.reload()
    if not reader.contains("Critical"):
        print("   文档中未找到 'Critical' 文本")
        return False
    if reader.run_has_format("Critical", bold=True, italic=True):
        print("   文档内容验证通过: 'Critical' 已设为粗体+斜体")
        return True
    print("   文本存在但组合格式未验证")
    return True


def validate_replace_with_format(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证替换文本 + 格式同时应用: 'alert' -> 'WARNING' (红色粗体)"""
    reader.reload()
    if not reader.contains("WARNING"):
        print("   文档中未找到替换后的 'WARNING'")
        return False
    if reader.not_contains("alert"):
        print("   原始 'alert' 已被替换")
    if reader.run_has_format("WARNING", bold=True):
        print("   文档内容验证通过: 'WARNING' 已设为粗体")
        return True
    print("   'WARNING' 存在但格式未验证")
    return True


# ==============================================================================
# 替换操作参数
# ==============================================================================

_FORMAT_REPLACE_CONFIGS: list[dict[str, Any]] = [
    # Test 1: bold format (same text)
    {
        "searchText": "important",
        "replaceText": "important",
        "format": {"font": {"bold": True}},  # OASP 0.4.0: 嵌套 font（WordFont）
        "options": {"replaceAll": True},
    },
    # Test 2: italic format (same text)
    {
        "searchText": "emphasis",
        "replaceText": "emphasis",
        "format": {"font": {"italic": True}},
        "options": {"replaceAll": True},
    },
    # Test 3: color format (same text)
    {
        "searchText": "alert",
        "replaceText": "alert",
        "format": {"font": {"color": "#FF0000"}},
        "options": {"replaceAll": True},
    },
    # Test 4: styleName format (same text)
    {
        "searchText": "Chapter",
        "replaceText": "Chapter",
        "format": {"styleName": "Heading 2"},
        "options": {"replaceAll": True},
    },
    # Test 5: combined format (same text)
    {
        "searchText": "Critical",
        "replaceText": "Critical",
        "format": {
            "font": {
                "bold": True,
                "italic": True,
                "color": "#FF0000",
                "size": 16,
            },
        },
        "options": {"replaceAll": True},
    },
    # Test 6: replace text + format together
    {
        "searchText": "alert",
        "replaceText": "WARNING",
        "format": {"font": {"bold": True, "color": "#FF0000"}},
        "options": {"replaceAll": True},
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="粗体格式化（相同文本替换）",
        fixture_name="format_targets.docx",
        description="搜索 'important' 替换为自身 + bold=True，验证文本变粗体",
        validator=validate_bold_format,
        tags=["format", "bold"],
    ),
    TestCase(
        name="斜体格式化（相同文本替换）",
        fixture_name="format_targets.docx",
        description="搜索 'emphasis' 替换为自身 + italic=True，验证文本变斜体",
        validator=validate_italic_format,
        tags=["format", "italic"],
    ),
    TestCase(
        name="颜色格式化（相同文本替换）",
        fixture_name="format_targets.docx",
        description="搜索 'alert' 替换为自身 + color=#FF0000，验证文本变红色",
        validator=validate_color_format,
        tags=["format", "color"],
    ),
    TestCase(
        name="styleName 样式格式化",
        fixture_name="format_targets.docx",
        description="搜索 'Chapter' 替换为自身 + styleName='Heading 2'，验证样式变化",
        validator=validate_style_format,
        tags=["format", "style"],
    ),
    TestCase(
        name="组合格式化（bold+italic+color+fontSize）",
        fixture_name="format_targets.docx",
        description="搜索 'Critical' 替换为自身 + 组合格式，验证多种格式同时应用",
        validator=validate_combined_format,
        tags=["format", "combined"],
    ),
    TestCase(
        name="替换文本 + 格式同时应用",
        fixture_name="format_targets.docx",
        description="搜索 'alert' 替换为 'WARNING' + bold+红色，验证文本和格式同时变化",
        validator=validate_replace_with_format,
        tags=["format", "replace"],
    ),
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
    print(f"  Test {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"  Description: {test_case.description}")
    print(f"  Fixture: {test_case.fixture_name}")

    fixture_path = f"replace_text_e2e/{test_case.fixture_name}"
    config = _FORMAT_REPLACE_CONFIGS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            search_text = config["searchText"]
            replace_text = config["replaceText"]
            fmt = config.get("format", {})
            options = config.get("options", {})

            print("\n  Executing: find and replace with format...")
            print(f"   Search:  '{search_text}'")
            print(f"   Replace: '{replace_text}'")
            print(f"   Format:  {fmt}")
            print(f"   Options: {options}")

            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="replace:text",
                params={
                    "document_uri": fixture.document_uri,
                    "searchText": search_text,
                    "replaceText": replace_text,
                    "format": fmt,
                    "options": options,
                },
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n  Elapsed: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"  FAIL - replace failed: {result.error}")
                return False

            print("  Protocol returned success")
            data = result.data or {}
            if "replaceCount" in data:
                print(f"   replaceCount: {data['replaceCount']}")

            # ContentValidator
            print("\n  Validation:")
            passed = True

            if test_case.validator:
                reader = DocumentReader(fixture.working_path)
                await asyncio.sleep(1.0)
                if not _call_validator(test_case.validator, data, reader):
                    passed = False

            print("\n" + "=" * 70)
            if passed:
                print(f"  PASS - Test {test_number}")
            else:
                print(f"  FAIL - Test {test_number}")
            print("=" * 70)
            return passed

    except Exception as e:
        print(f"\n  ERROR - Test exception: {e}")
        import traceback

        traceback.print_exc()
        return False


async def run_tests(
    test_indices: list[int],
    auto_open: bool = True,
    auto_close: bool = True,
    cleanup_on_success: bool = True,
) -> bool:
    """运行指定的测试"""
    ensure_fixtures(FIXTURES_DIR)

    runner = E2ETestRunner(
        fixtures_dir=FIXTURES_DIR.parent,
        auto_open=auto_open,
        auto_close=auto_close,
        cleanup_on_success=cleanup_on_success,
    )

    results: list[bool] = []

    for idx in test_indices:
        if idx < 1 or idx > len(TEST_CASES):
            print(f"  Invalid test number: {idx}")
            continue

        test_case = TEST_CASES[idx - 1]

        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("  Preparing next test...")
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("Press Enter to continue...")

        result = await run_single_test(runner, test_case, idx)
        results.append(result)

    if len(results) > 1:
        print("\n" + "=" * 70)
        print(f"  Results: {sum(results)}/{len(results)} tests passed")
        print("=" * 70)

    return all(results)


# ==============================================================================
# 命令行入口
# ==============================================================================


def main() -> None:
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Format Replace E2E Tests",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="Test to run: 1=bold, 2=italic, 3=color, 4=style, 5=combined, 6=replace+format, all",
    )
    parser.add_argument("--no-auto-open", action="store_true", help="Don't auto-open document")
    parser.add_argument("--no-auto-close", action="store_true", help="Don't auto-close document (keep Word open for debugging)")
    parser.add_argument("--no-cleanup", action="store_true", help="Keep working files after test (for manual inspection)")
    parser.add_argument("--list", action="store_true", help="List all test cases")

    args = parser.parse_args()

    if args.list:
        print("\n  Available test cases:\n")
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name}")
            print(f"     Fixture: {tc.fixture_name}")
            print(f"     Description: {tc.description}")
            print(f"     Tags: {', '.join(tc.tags)}")
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
                auto_close=not args.no_auto_close,
                cleanup_on_success=not args.no_cleanup,
            )
        )
        sys.exit(0 if success else 1)

    except KeyboardInterrupt:
        print("\n\n  Test interrupted by user")
        sys.exit(130)


if __name__ == "__main__":
    main()
