"""
Basic Select Text E2E Tests (自动化版本)

测试基本的文本选择功能。

测试场景:
1. 简单选中文本
2. 选择第N个匹配
3. 不区分大小写
4. 全字匹配

运行方式:
    uv run python manual_tests/select_text_e2e/test_basic_select.py --test 1
    uv run python manual_tests/select_text_e2e/test_basic_select.py --test all
    uv run python manual_tests/select_text_e2e/test_basic_select.py --list
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

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "select_text_e2e"

# ==============================================================================
# Validator 函数
# ==============================================================================


def validate_simple_select(data: dict[str, Any]) -> bool:
    """验证简单选中: 找到并选中 'Hello World'"""
    match_count = data.get("matchCount", 0)
    selected_text = data.get("selectedText")
    if match_count > 0 and selected_text == "Hello World":
        print(f"   ✅ 匹配数: {match_count}, 选中文本: '{selected_text}'")
        return True
    print(f"   ❌ matchCount={match_count}, selectedText='{selected_text}' (预期 'Hello World')")
    return False


def validate_select_nth(data: dict[str, Any]) -> bool:
    """验证选择第N个匹配: 至少有 2 个 'test' 匹配"""
    match_count = data.get("matchCount", 0)
    if match_count >= 2:
        print(f"   ✅ 匹配数: {match_count} (>= 2), 成功选中第 2 个匹配")
        return True
    print(f"   ❌ matchCount={match_count} (预期 >= 2)")
    return False


def validate_case_insensitive(data: dict[str, Any]) -> bool:
    """验证不区分大小写: 搜索 'hello' 能匹配到"""
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ 不区分大小写匹配成功, matchCount={match_count}")
        return True
    print(f"   ❌ matchCount={match_count} (预期 > 0)")
    return False


def validate_whole_word(data: dict[str, Any]) -> bool:
    """验证全字匹配: 搜索 'test' 只匹配完整单词"""
    match_count = data.get("matchCount", 0)
    selected_text = data.get("selectedText")
    if match_count > 0 and selected_text == "test":
        print(f"   ✅ 全字匹配成功, matchCount={match_count}, selectedText='{selected_text}'")
        return True
    print(f"   ❌ matchCount={match_count}, selectedText='{selected_text}' (预期 'test')")
    return False


# ==============================================================================
# 选择操作参数
# ==============================================================================

_SELECT_CONFIGS: list[dict[str, Any]] = [
    # Test 1: 简单选中
    {
        "search_text": "Hello World",
    },
    # Test 2: 选择第N个匹配
    {
        "search_text": "test",
        "select_index": 2,
    },
    # Test 3: 不区分大小写
    {
        "search_text": "hello",
        "search_options": {"matchCase": False},
    },
    # Test 4: 全字匹配
    {
        "search_text": "test",
        "search_options": {"matchWholeWord": True},
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="简单选中文本",
        fixture_name="simple.docx",
        description="搜索 'Hello World' 并选中，验证 matchCount > 0 且 selectedText 正确",
        validator=validate_simple_select,
        tags=["basic", "select"],
    ),
    TestCase(
        name="选择第N个匹配",
        fixture_name="simple.docx",
        description="搜索 'test' 并选中第 2 个匹配，验证 matchCount >= 2",
        validator=validate_select_nth,
        tags=["basic", "select_index"],
    ),
    TestCase(
        name="不区分大小写",
        fixture_name="simple.docx",
        description="搜索 'hello' (matchCase=False)，验证能匹配到大小写不同的文本",
        validator=validate_case_insensitive,
        tags=["basic", "case_insensitive"],
    ),
    TestCase(
        name="全字匹配",
        fixture_name="simple.docx",
        description="搜索 'test' (matchWholeWord=True)，验证只匹配完整单词",
        validator=validate_whole_word,
        tags=["basic", "whole_word"],
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

    fixture_path = f"select_text_e2e/{test_case.fixture_name}"
    config = _SELECT_CONFIGS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            search_text = config["search_text"]
            selection_mode = config.get("selection_mode", "select")
            select_index = config.get("select_index", 1)
            search_options = config.get("search_options")

            print(f"\n📝 执行: 选择文本...")
            print(f"   搜索: '{search_text}'")
            print(f"   模式: {selection_mode}, 索引: {select_index}")
            if search_options:
                print(f"   搜索选项: {search_options}")

            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="select:text",
                params={
                    "document_uri": fixture.document_uri,
                    "search_text": search_text,
                    "selection_mode": selection_mode,
                    "select_index": select_index,
                    **({"search_options": search_options} if search_options else {}),
                },
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"❌ 选择失败: {result.error}")
                return False

            print("✅ 协议返回成功")
            data = result.data or {}
            if "matchCount" in data:
                print(f"   匹配数: {data['matchCount']}")
            if "selectedText" in data:
                print(f"   选中文本: '{data['selectedText']}'")

            # DataValidator 验证
            print("\n📊 验证结果:")
            passed = True

            if test_case.validator:
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
        description="Basic Select Text E2E Tests (自动化版本)",
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
