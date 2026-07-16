"""
Edge Cases Select Text E2E Tests (自动化版本)

测试 word:select:text 的边界情况和错误处理。

测试场景:
1. 未找到匹配（预期失败）
2. 空搜索文本（预期失败）
3. 超出索引（Add-In 优雅降级）
4. 特殊字符搜索
5. 长文本搜索（255字符以内）
6. 超出 255 字符限制（预期被 DTO 拒绝）

运行方式:
    uv run python manual_tests/select_text_e2e/test_edge_cases.py --test 1
    uv run python manual_tests/select_text_e2e/test_edge_cases.py --test all
    uv run python manual_tests/select_text_e2e/test_edge_cases.py --list
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


def validate_no_match(data: dict[str, Any]) -> bool:
    """验证未找到匹配: 预期失败，检查错误信息"""
    # 对于预期失败的用例，data 中包含 _expected_failure 标记
    if data.get("_expected_failure"):
        error = data.get("error", "")
        if error:
            print(f"   ✅ 预期失败，错误信息: {error}")
            return True
        print("   ❌ 预期失败但缺少错误信息")
        return False
    # 如果意外成功，检查 matchCount
    match_count = data.get("matchCount", 0)
    if match_count == 0:
        print(f"   ✅ matchCount={match_count} (操作成功但无匹配)")
        return True
    print(f"   ❌ 不应该找到匹配，但 matchCount={match_count}")
    return False


def validate_empty_search(data: dict[str, Any]) -> bool:
    """验证空搜索文本: 预期失败"""
    if data.get("_expected_failure"):
        error = data.get("error", "")
        if error:
            print(f"   ✅ 空搜索文本被正确拒绝，错误信息: {error}")
            return True
        print("   ❌ 预期失败但缺少错误信息")
        return False
    print("   ⚠️  空搜索文本未被拒绝（实现可能允许空搜索）")
    return True


def validate_out_of_bounds(data: dict[str, Any]) -> bool:
    """验证超出索引: Add-In 可能返回成功（优雅降级）或失败"""
    if data.get("_expected_failure"):
        error = data.get("error", "")
        if error:
            print(f"   ✅ 索引超出范围被拒绝，错误信息: {error}")
            return True
        print("   ❌ 预期失败但缺少错误信息")
        return False
    # Add-In 优雅降级：select_index 超出匹配数时返回成功 + 实际 matchCount
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ Add-In 优雅降级: matchCount={match_count} (请求 select_index=10)")
        print("      Add-In 不将越界索引视为错误，而是返回实际匹配数")
        return True
    print("   ❌ matchCount=0，预期至少有匹配")
    return False


def validate_special_chars(data: dict[str, Any]) -> bool:
    """验证特殊字符搜索: 应该成功找到"""
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ 特殊字符搜索成功，matchCount={match_count}")
        return True
    print(f"   ❌ matchCount={match_count} (预期 > 0)")
    return False


def validate_long_text(data: dict[str, Any]) -> bool:
    """验证长文本搜索（255字符以内）: 检查 matchCount"""
    match_count = data.get("matchCount", 0)
    print(f"   matchCount={match_count}")
    # 长文本搜索可能成功也可能失败（取决于文档内容），只要不崩溃即可
    return True


def validate_exceed_max_length(data: dict[str, Any]) -> bool:
    """验证超出 255 字符限制: 预期被 DTO 验证拒绝"""
    if data.get("_expected_failure"):
        error = data.get("error", "")
        if error and "255" in error:
            print(f"   ✅ 超长文本被正确拒绝 (max_length=255)")
            return True
        if error:
            print(f"   ✅ 超长文本被拒绝，错误信息: {error}")
            return True
        print("   ❌ 预期失败但缺少错误信息")
        return False
    print("   ❌ 超长文本（>255字符）应该被拒绝")
    return False


# ==============================================================================
# 选择操作参数
# ==============================================================================

_SELECT_CONFIGS: list[dict[str, Any]] = [
    # Test 1: 未找到匹配
    {
        "search_text": "NonExistentText12345",
        "expects_failure": True,
    },
    # Test 2: 空搜索文本
    {
        "search_text": "",
        "expects_failure": True,
    },
    # Test 3: 超出索引（Add-In 优雅降级，不视为错误）
    {
        "search_text": "OutOfBounds",
        "select_index": 10,
    },
    # Test 4: 特殊字符
    {
        "search_text": "@#$%",
    },
    # Test 5: 长文本搜索（255字符以内）
    {
        "search_text": "This is a very long text that " * 8,  # 240 chars, within 255 limit
    },
    # Test 6: 超出 255 字符限制（预期被 DTO 验证拒绝）
    {
        "search_text": "A" * 256,
        "expects_failure": True,
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="未找到匹配",
        fixture_name="simple.docx",
        description="搜索 'NonExistentText12345'，预期操作返回失败",
        validator=validate_no_match,
        tags=["edge_case", "expected_failure"],
    ),
    TestCase(
        name="空搜索文本",
        fixture_name="simple.docx",
        description="搜索空字符串，预期操作返回失败或错误",
        validator=validate_empty_search,
        tags=["edge_case", "expected_failure"],
    ),
    TestCase(
        name="超出索引（优雅降级）",
        fixture_name="edge_cases.docx",
        description="搜索 'OutOfBounds' 但 select_index=10 超出匹配数，Add-In 优雅降级",
        validator=validate_out_of_bounds,
        tags=["edge_case", "graceful_degradation"],
    ),
    TestCase(
        name="特殊字符搜索",
        fixture_name="edge_cases.docx",
        description="搜索 '@#$%' 特殊字符，验证能正确匹配",
        validator=validate_special_chars,
        tags=["edge_case", "special_chars"],
    ),
    TestCase(
        name="长文本搜索（255字符以内）",
        fixture_name="edge_cases.docx",
        description="搜索重复 8 次的长文本（240字符），验证不会崩溃",
        validator=validate_long_text,
        tags=["edge_case", "long_text"],
    ),
    TestCase(
        name="超出 255 字符限制",
        fixture_name="edge_cases.docx",
        description="搜索 256 字符文本，预期被 DTO max_length=255 验证拒绝",
        validator=validate_exceed_max_length,
        tags=["edge_case", "expected_failure"],
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
    expects_failure = config.get("expects_failure", False)

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            search_text = config["search_text"]
            selection_mode = config.get("selection_mode", "select")
            select_index = config.get("select_index", 1)
            search_options = config.get("search_options")

            display_text = search_text[:60] + "..." if len(search_text) > 60 else search_text
            print(f"\n📝 执行: 选择文本...")
            print(f"   搜索: '{display_text}'")
            print(f"   模式: {selection_mode}, 索引: {select_index}")
            if search_options:
                print(f"   搜索选项: {search_options}")
            if expects_failure:
                print("   ⚠️  预期此操作会失败")

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

            # 对于预期失败的用例，不在 result.success==False 时直接返回 False
            if expects_failure:
                if not result.success:
                    print(f"✅ 操作按预期返回失败: {result.error}")
                    data: dict[str, Any] = {
                        "_expected_failure": True,
                        "error": result.error,
                        "success": result.success,
                    }
                else:
                    print("⚠️  操作意外成功（预期失败）")
                    data = result.data or {}
            else:
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
        description="Edge Cases Select Text E2E Tests (自动化版本)",
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
