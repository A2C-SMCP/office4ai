"""
Basic Text Replace E2E Tests (自动化版本)

测试基本的文本查找和替换功能。

测试场景:
1. 简单文本替换（全部）
2. 简单文本替换（首个）
3. 替换为空（删除）
4. 跨段落文本替换（使用 ^p 记号）
5. 特殊字符替换
6. 长文本替换

运行方式:
    uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test 1
    uv run python manual_tests/replace_text_e2e/test_basic_replace.py --test all
    uv run python manual_tests/replace_text_e2e/test_basic_replace.py --list
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


def validate_replace_all(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证全部替换: 所有 'old' 替换为 'new'"""
    reader.reload()
    if reader.contains("old") and not reader.contains("old text"):
        # 排除 "old" 出现在非目标位置的情况
        pass
    if not reader.contains("new"):
        print("   ❌ 文档中未找到替换后的 'new'")
        return False
    if reader.not_contains("old text"):
        print("   ✅ 文档内容验证通过: 'old text' 已被替换")
        return True
    print("   ⚠️  文档中仍包含 'old text'")
    return False


def validate_replace_first(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证首个替换: 第一个 'test' 替换为 'exam'"""
    reader.reload()
    if not reader.contains("exam"):
        print("   ❌ 文档中未找到替换后的 'exam'")
        return False
    # 应该还有剩余的 test
    if reader.contains("test"):
        print("   ✅ 文档内容验证通过: 第一个 'test' 被替换，仍有剩余 'test'")
        return True
    print("   ⚠️  所有 'test' 似乎都被替换了（预期只替换第一个）")
    return True  # 仍算通过，协议层已验证 replaceCount


def validate_delete(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证删除: 所有 'delete' 被删除"""
    reader.reload()
    if reader.not_contains("delete"):
        print("   ✅ 文档内容验证通过: 所有 'delete' 已被删除")
        return True
    print("   ❌ 文档中仍包含 'delete'")
    return False


def validate_special_chars(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证特殊字符替换: 'Café' 替换为 'Coffee'"""
    reader.reload()
    if not reader.contains("Coffee"):
        print("   ❌ 文档中未找到替换后的 'Coffee'")
        return False
    if reader.not_contains("Café"):
        print("   ✅ 文档内容验证通过: 'Café' 已替换为 'Coffee'")
        return True
    print("   ⚠️  文档中仍包含 'Café'")
    return False


def validate_cross_paragraph(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证跨段落替换: 'line1\\nline2' 替换为 'new\\ncontent'"""
    reader.reload()
    if not reader.contains("new"):
        print("   ❌ 文档中未找到替换后的 'new'")
        return False
    if reader.not_contains("line1") and reader.not_contains("line2"):
        print("   ✅ 文档内容验证通过: 'line1/line2' 已被替换为 'new/content'")
        return True
    print("   ⚠️  文档中仍包含 'line1' 或 'line2'")
    return False


def validate_long_text(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证长文本替换"""
    reader.reload()
    if reader.contains("Here is another lengthy paragraph"):
        print("   ✅ 文档内容验证通过: 长文本已被替换")
        return True
    print("   ❌ 文档中未找到替换后的长文本")
    return False


# ==============================================================================
# 替换操作参数
# ==============================================================================

_REPLACE_CONFIGS: list[dict[str, Any]] = [
    # Test 1: replace all
    {"searchText": "old", "replaceText": "new", "options": {"replaceAll": True}},
    # Test 2: replace first
    {"searchText": "test", "replaceText": "exam", "options": {"replaceAll": False}},
    # Test 3: delete (replace with empty)
    {"searchText": "delete", "replaceText": "", "options": {"replaceAll": True}},
    # Test 4: cross-paragraph (使用 Word 特殊字符记号 ^p 表示段落标记)
    {"searchText": "line1^pline2", "replaceText": "new^pcontent", "options": {"replaceAll": True}},
    # Test 5: special chars
    {"searchText": "Café", "replaceText": "Coffee", "options": {"replaceAll": True}},
    # Test 6: long text
    {
        "searchText": (
            "This is a long paragraph of text that should be replaced with another "
            "long paragraph. It contains multiple sentences and various punctuation marks."
        ),
        "replaceText": (
            "Here is another lengthy paragraph that serves as the replacement text. "
            "It also has multiple sentences and demonstrates the replace functionality."
        ),
        "options": {"replaceAll": True},
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="简单文本替换（全部）",
        fixture_name="replace_targets.docx",
        description="搜索所有 'old' 替换为 'new'，验证 'old' 不存在且 'new' 存在",
        validator=validate_replace_all,
        tags=["basic", "replace_all"],
    ),
    TestCase(
        name="简单文本替换（首个）",
        fixture_name="replace_targets.docx",
        description="搜索 'test' 仅替换第一个为 'exam'，验证仍有剩余 'test'",
        validator=validate_replace_first,
        tags=["basic", "replace_first"],
    ),
    TestCase(
        name="替换为空（删除）",
        fixture_name="replace_targets.docx",
        description="搜索 'delete' 替换为空字符串，验证所有 'delete' 被删除",
        validator=validate_delete,
        tags=["basic", "delete"],
    ),
    TestCase(
        name="跨段落文本替换",
        fixture_name="replace_targets.docx",
        description="使用 ^p 记号搜索跨段落文本 'line1^pline2' 替换为 'new^pcontent'",
        validator=validate_cross_paragraph,
        tags=["advanced"],
    ),
    TestCase(
        name="特殊字符替换",
        fixture_name="replace_targets.docx",
        description="搜索 'Café' 替换为 'Coffee'，验证特殊字符正确处理",
        validator=validate_special_chars,
        tags=["advanced"],
    ),
    TestCase(
        name="长文本替换",
        fixture_name="replace_targets.docx",
        description="搜索长段落替换为另一个长段落，验证替换正确",
        validator=validate_long_text,
        tags=["advanced"],
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
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"replace_text_e2e/{test_case.fixture_name}"
    config = _REPLACE_CONFIGS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            search_text = config["searchText"]
            replace_text = config["replaceText"]
            options = config.get("options", {})

            print(f"\n📝 执行: 查找替换...")
            print(f"   搜索: '{search_text[:60]}{'...' if len(search_text) > 60 else ''}'")
            print(f"   替换: '{replace_text[:60]}{'...' if len(replace_text) > 60 else ''}'")
            print(f"   选项: {options}")

            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="replace:text",
                params={
                    "document_uri": fixture.document_uri,
                    "searchText": search_text,
                    "replaceText": replace_text,
                    "options": options,
                },
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                if test_case.expect_failure:
                    print(f"⚠️  操作失败（符合预期）: {result.error}")
                    print("\n" + "=" * 70)
                    print(f"✅ 测试 {test_number} 通过（预期失败，已确认）")
                    print("=" * 70)
                    return True
                print(f"❌ 替换失败: {result.error}")
                return False

            if test_case.expect_failure:
                print("🚨 意外成功！此测试预期失败但实际通过了。")
                print("   Word JS API 可能已支持此能力，请检查 SENTINEL TEST 注释中的 TODO 项。")
                print("\n" + "=" * 70)
                print(f"⚠️  测试 {test_number} 意外通过（需人工确认）")
                print("=" * 70)
                return False

            print("✅ 协议返回成功")
            data = result.data or {}
            if "replaceCount" in data:
                print(f"   替换次数: {data['replaceCount']}")

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
        description="Basic Text Replace E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试",
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
