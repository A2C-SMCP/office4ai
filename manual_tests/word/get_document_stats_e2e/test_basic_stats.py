"""
Basic Get Document Stats Test (自动化版本)

测试文档统计信息获取功能。

特性：
- 自动复制测试文档
- 自动打开 Word
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. 空白文档统计 - 获取空白文档的统计信息
2. 简单文档统计 - 获取包含简单文本的文档统计
3. 复杂文档统计 - 获取包含多种元素的文档统计
4. 大文档统计 - 获取包含大量内容的文档统计（性能测试）

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_document_stats_e2e/test_basic_stats.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_document_stats_e2e/test_basic_stats.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_document_stats_e2e/test_basic_stats.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_document_stats_e2e/test_basic_stats.py --test 1 --always-cleanup
"""

import asyncio
import sys
import time
from pathlib import Path

from manual_tests.e2e_base import (
    DocumentReader,
    E2ETestRunner,
    ExpectedStats,
    TestCase,
    _call_validator,
    ensure_fixtures,
)
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

# 夹具目录
FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "get_document_stats_e2e"

# 测试用例定义
TEST_CASES: list[TestCase] = [
    TestCase(
        name="空白文档统计",
        fixture_name="empty.docx",
        description="空白文档应该有 0 字、0 字符、1 段落（Word 默认有一个空段落）",
        expected=ExpectedStats(
            word_count=0,
            character_count=0,
            paragraph_count=1,  # Word 空文档默认有 1 个段落
            paragraph_count_tolerance=1,  # 允许 0 或 1
        ),
        tags=["basic"],
    ),
    TestCase(
        name="简单文档统计",
        fixture_name="simple.docx",
        description="简单文本文档应该正确统计字数、字符数和段落数",
        expected=ExpectedStats(
            word_count=None,  # 不精确验证，只检查 > 0
            character_count=None,
            paragraph_count=None,
        ),
        # 自定义验证：确保所有值都 > 0
        validator=lambda data: (
            data.get("wordCount", 0) > 0
            and data.get("characterCount", 0) > 0
            and data.get("paragraphCount", 0) > 0
        ),
        tags=["basic"],
    ),
    TestCase(
        name="复杂文档统计",
        fixture_name="complex.docx",
        description="包含表格、列表的文档应该正确统计",
        expected=ExpectedStats(
            word_count=None,
            character_count=None,
            paragraph_count=None,
        ),
        validator=lambda data: (
            data.get("wordCount", 0) > 0
            and data.get("characterCount", 0) > 0
            and data.get("paragraphCount", 0) > 0
        ),
        tags=["advanced"],
    ),
    TestCase(
        name="大文档统计",
        fixture_name="large.docx",
        description="大文档应该能快速返回统计结果（性能测试）",
        expected=ExpectedStats(
            word_count=None,
            character_count=None,
            paragraph_count=None,
        ),
        validator=lambda data: (
            data.get("wordCount", 0) > 1000  # 大文档应该有较多字数
            and data.get("characterCount", 0) > 5000
        ),
        tags=["performance"],
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
    """
    执行单个测试用例

    Args:
        runner: E2E 测试运行器
        test_case: 测试用例
        test_number: 测试编号

    Returns:
        是否通过
    """
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"get_document_stats_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 执行获取文档统计动作
            print("\n📝 执行: 获取文档统计...")
            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="get:documentStats",
                params={"document_uri": fixture.document_uri},
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            # 显示结果
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"❌ 获取失败: {result.error}")
                return False

            print("✅ 获取成功")
            data = result.data or {}
            print(f"   字数: {data.get('wordCount', 'N/A')}")
            print(f"   字符数: {data.get('characterCount', 'N/A')}")
            print(f"   段落数: {data.get('paragraphCount', 'N/A')}")

            # 验证结果
            print("\n📊 验证结果:")
            passed = True

            # 预期值验证
            if test_case.expected:
                verified, messages = runner.verify_stats(data, test_case.expected)
                for msg in messages:
                    print(f"   {msg}")
                if not verified:
                    passed = False

            # 自定义验证
            if test_case.validator:
                # 创建文档读取器（用于双重验证）
                reader = DocumentReader(fixture.working_path)
                if _call_validator(test_case.validator, data, reader):
                    print("   ✅ 自定义验证通过")
                else:
                    print("   ❌ 自定义验证失败")
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
    """
    运行指定的测试

    Args:
        test_indices: 要运行的测试索引列表（1-based）
        auto_open: 是否自动打开文档
        cleanup_on_success: 成功后是否清理

    Returns:
        是否全部通过
    """
    # 确保夹具存在
    ensure_fixtures(FIXTURES_DIR)

    # 创建运行器
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

        # 如果运行多个测试，等待用户确认（给时间切换文档）
        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("⏳ 准备下一个测试...")
            if auto_open:
                # 自动模式：等待几秒
                await asyncio.sleep(2.0)
            else:
                # 手动模式：等待用户按回车
                input("按回车继续...")

        result = await run_single_test(runner, test_case, idx)
        results.append(result)

    # 汇总结果
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
        description="Word Get Document Stats E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（空白文档）
  python test_basic_stats.py --test 1

  # 运行所有测试
  python test_basic_stats.py --test all

  # 手动打开文档模式
  python test_basic_stats.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_basic_stats.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="要运行的测试: 1=空白, 2=简单, 3=复杂, 4=大文档, all=全部",
    )

    parser.add_argument(
        "--no-auto-open",
        action="store_true",
        help="不自动打开文档（需要手动打开）",
    )

    parser.add_argument(
        "--always-cleanup",
        action="store_true",
        help="无论成功失败都清理测试文件",
    )

    parser.add_argument(
        "--list",
        action="store_true",
        help="列出所有测试用例",
    )

    args = parser.parse_args()

    # 列出测试
    if args.list:
        print("\n📋 可用测试用例:\n")
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name}")
            print(f"     夹具: {tc.fixture_name}")
            print(f"     描述: {tc.description}")
            print(f"     标签: {', '.join(tc.tags)}")
            print()
        return

    # 解析测试索引
    if args.test == "all":
        test_indices = list(range(1, len(TEST_CASES) + 1))
    else:
        test_indices = [int(args.test)]

    # 运行测试
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
