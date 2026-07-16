"""
Basic Insert Text E2E Tests (自动化版本)

测试基本的文本插入功能（使用 location="End" 避免光标定位问题）。

测试场景:
1. 简单文本插入
2. 多行文本插入
3. 特殊字符插入
4. 长文本插入

运行方式:
    uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test 1
    uv run python manual_tests/insert_text_e2e/test_basic_insert.py --test all
    uv run python manual_tests/insert_text_e2e/test_basic_insert.py --list
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

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "insert_text_e2e"

# ==============================================================================
# Validator 函数
# ==============================================================================


def validate_simple_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证简单文本插入"""
    reader.reload()
    if not reader.contains("Hello World"):
        print("   ❌ 文档中未找到 'Hello World'")
        return False
    print("   ✅ 文档内容验证通过: 包含 'Hello World'")
    return True


def validate_multiline_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证多行文本插入"""
    reader.reload()
    for line in ["第一行文本", "第二行文本", "第三行文本"]:
        if not reader.contains(line):
            print(f"   ❌ 文档中未找到 '{line}'")
            return False
    print("   ✅ 文档内容验证通过: 包含所有行")
    return True


def validate_special_chars_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证特殊字符插入"""
    reader.reload()
    if not reader.contains("@#$%^&*()"):
        print("   ❌ 文档中未找到特殊字符")
        return False
    print("   ✅ 文档内容验证通过: 包含特殊字符")
    return True


def validate_long_text_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证长文本插入"""
    reader.reload()
    if not reader.contains("这是一段较长的文本"):
        print("   ❌ 文档中未找到长文本")
        return False
    if not reader.contains("确保插入长文本不会导致系统卡顿或超时"):
        print("   ❌ 文档中未找到长文本尾部内容")
        return False
    print("   ✅ 文档内容验证通过: 长文本完整插入")
    return True


# ==============================================================================
# 测试数据
# ==============================================================================

SIMPLE_TEXT = "Hello World"

MULTILINE_TEXT = """第一行文本
第二行文本
第三行文本"""

SPECIAL_TEXT = "特殊字符测试: @#$%^&*()_+-=[]{}|;':\",./<>?~`"

LONG_TEXT = (
    "这是一段较长的文本，用于测试 Word Add-In 处理较长内容的能力。"
    "这段文本包含了多个句子，每个句子都测试不同的字符和标点符号。\n"
    "在插入这段文本后，我们应该验证：\n"
    "1. 文本是否完整插入\n"
    "2. 格式是否保持正确\n"
    "3. 是否有乱码或丢失字符\n\n"
    "此外，我们还需要测试性能，确保插入长文本不会导致系统卡顿或超时。"
    "这个测试对于确保用户体验非常重要，因为在实际使用中，用户可能会插入大段文本。"
)

# 测试用例定义
TEST_CASES: list[TestCase] = [
    TestCase(
        name="简单文本插入",
        fixture_name="empty.docx",
        description="在空文档末尾插入 'Hello World'，验证文档包含该文本",
        validator=validate_simple_insert,
        tags=["basic"],
    ),
    TestCase(
        name="多行文本插入",
        fixture_name="empty.docx",
        description="在空文档末尾插入多行文本，验证每行都存在",
        validator=validate_multiline_insert,
        tags=["basic"],
    ),
    TestCase(
        name="特殊字符插入",
        fixture_name="empty.docx",
        description="在空文档末尾插入包含特殊字符的文本，验证字符完整",
        validator=validate_special_chars_insert,
        tags=["basic"],
    ),
    TestCase(
        name="长文本插入",
        fixture_name="empty.docx",
        description="在空文档末尾插入长文本（含多段落），验证完整性",
        validator=validate_long_text_insert,
        tags=["basic"],
    ),
]

# 每个测试用例对应的插入文本
_INSERT_TEXTS = [SIMPLE_TEXT, MULTILINE_TEXT, SPECIAL_TEXT, LONG_TEXT]


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

    fixture_path = f"insert_text_e2e/{test_case.fixture_name}"
    insert_text = _INSERT_TEXTS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 执行插入
            print(f"\n📝 执行: 插入文本 (location=End)...")
            print(f"   文本预览: '{insert_text[:60]}{'...' if len(insert_text) > 60 else ''}'")
            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="insert:text",
                params={
                    "document_uri": fixture.document_uri,
                    "text": insert_text,
                    "location": "End",
                },
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"❌ 插入失败: {result.error}")
                return False

            print("✅ 协议返回成功")
            data = result.data or {}

            # ContentValidator 双重验证
            print("\n📊 验证结果:")
            passed = True

            if test_case.validator:
                # DocumentReader 的 reload() 会自动通过 AppleScript 保存文档
                await asyncio.sleep(0.5)
                reader = DocumentReader(fixture.working_path)
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
        description="Basic Insert Text E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=简单, 2=多行, 3=特殊字符, 4=长文本, all=全部",
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
