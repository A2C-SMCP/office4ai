"""
Basic Text Replace Selection E2E Tests (自动化版本)

测试基本的选区文本替换功能（word:replace:selection）。

注意：这些测试需要用户手动在 Word 中选中文本后按 Enter 继续。

测试场景:
1. 替换为纯文本 - 选中文本后替换为 "Hello World"
2. 替换为多行文本 - 选中文本后替换为多行内容
3. 替换为特殊字符 - 选中文本后替换为包含特殊字符的文本
4. 替换为长文本 - 选中文本后替换为重复 50 次的长文本

运行方式:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test 1

    # 运行所有测试
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/replace_selection_e2e/test_text_replace.py --test 1 --always-cleanup
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


def validate_simple_text(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证纯文本替换"""
    reader.reload()
    if not reader.contains("Hello World"):
        print("   ❌ 文档中未找到 'Hello World'")
        return False
    print("   ✅ 文档内容验证通过: 包含 'Hello World'")
    return True


def validate_multiline_text(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证多行文本替换"""
    reader.reload()
    lines = ["第一行替换内容", "第二行替换内容", "第三行替换内容"]
    for line in lines:
        if not reader.contains(line):
            print(f"   ❌ 文档中未找到 '{line}'")
            return False
    print("   ✅ 文档内容验证通过: 所有行都存在")
    return True


def validate_special_chars(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证特殊字符替换"""
    reader.reload()
    special_text = "特殊字符测试：中文、English、123、!@#$%、符号「」【】"
    if not reader.contains(special_text):
        print(f"   ❌ 文档中未找到特殊字符文本")
        return False
    print("   ✅ 文档内容验证通过: 特殊字符正确")
    return True


def validate_long_text(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证长文本替换"""
    reader.reload()
    long_text = "这是一段很长的测试文本。" * 50
    if not reader.contains(long_text):
        print("   ❌ 文档中未找到长文本")
        return False
    print("   ✅ 文档内容验证通过: 长文本完整存在")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[TestCase] = [
    TestCase(
        name="替换为纯文本",
        fixture_name="simple.docx",
        description="选中文本后替换为 'Hello World'，验证文档包含替换后的文本",
        validator=validate_simple_text,
        tags=["text", "basic"],
    ),
    TestCase(
        name="替换为多行文本",
        fixture_name="simple.docx",
        description="选中文本后替换为多行内容，验证每一行都存在",
        validator=validate_multiline_text,
        tags=["text", "multiline"],
    ),
    TestCase(
        name="替换为特殊字符",
        fixture_name="simple.docx",
        description="选中文本后替换为包含中英文、数字、符号的混合文本",
        validator=validate_special_chars,
        tags=["text", "special"],
    ),
    TestCase(
        name="替换为长文本",
        fixture_name="simple.docx",
        description="选中文本后替换为重复 50 次的长文本，验证完整性",
        validator=validate_long_text,
        tags=["text", "performance"],
    ),
]

# 每个测试用例的替换内容
_REPLACE_CONTENTS: list[dict[str, Any]] = [
    # Test 1: 纯文本
    {"text": "Hello World"},
    # Test 2: 多行文本
    {"text": "第一行替换内容\n第二行替换内容\n第三行替换内容"},
    # Test 3: 特殊字符
    {"text": "特殊字符测试：中文、English、123、!@#$%、符号「」【】"},
    # Test 4: 长文本
    {"text": "这是一段很长的测试文本。" * 50},
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
        description="Replace Selection Text E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（纯文本替换）
  python test_text_replace.py --test 1

  # 运行所有测试
  python test_text_replace.py --test all

  # 手动打开文档模式
  python test_text_replace.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_text_replace.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=纯文本, 2=多行, 3=特殊字符, 4=长文本, all=全部",
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
