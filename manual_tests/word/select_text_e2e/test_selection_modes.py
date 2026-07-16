"""
Selection Modes Select Text E2E Tests (自动化版本)

测试 word:select:text 的选择模式功能。

测试场景:
1. select 模式（高亮选区）
2. start 模式（光标定位到开头）
3. end 模式（光标定位到结尾）
4. 模式切换（连续执行 select/start/end）

运行方式:
    uv run python manual_tests/select_text_e2e/test_selection_modes.py --test 1
    uv run python manual_tests/select_text_e2e/test_selection_modes.py --test all
    uv run python manual_tests/select_text_e2e/test_selection_modes.py --list
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


def validate_select_mode(data: dict[str, Any]) -> bool:
    """验证 select 模式: 文本应该被高亮选中"""
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ select 模式成功，matchCount={match_count}")
        return True
    print(f"   ❌ matchCount={match_count} (预期 > 0)")
    return False


def validate_start_mode(data: dict[str, Any]) -> bool:
    """验证 start 模式: 光标应定位到匹配文本开头"""
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ start 模式成功，matchCount={match_count}")
        return True
    print(f"   ❌ matchCount={match_count} (预期 > 0)")
    return False


def validate_end_mode(data: dict[str, Any]) -> bool:
    """验证 end 模式: 光标应定位到匹配文本结尾"""
    match_count = data.get("matchCount", 0)
    if match_count > 0:
        print(f"   ✅ end 模式成功，matchCount={match_count}")
        return True
    print(f"   ❌ matchCount={match_count} (预期 > 0)")
    return False


def validate_mode_switching(data: dict[str, Any]) -> bool:
    """验证模式切换: 三种模式都应该成功执行"""
    # 自定义执行器会将三次执行的结果合并到 data 中
    results = data.get("_mode_results", {})
    all_passed = True

    for mode in ["select", "start", "end"]:
        mode_data = results.get(mode, {})
        success = mode_data.get("success", False)
        match_count = mode_data.get("matchCount", 0)
        if success and match_count > 0:
            print(f"   ✅ {mode} 模式: matchCount={match_count}")
        else:
            print(f"   ❌ {mode} 模式失败: success={success}, matchCount={match_count}")
            all_passed = False

    return all_passed


# ==============================================================================
# 选择操作参数
# ==============================================================================

_SELECT_CONFIGS: list[dict[str, Any]] = [
    # Test 1: select 模式
    {
        "search_text": "Selection Test",
        "selection_mode": "select",
    },
    # Test 2: start 模式
    {
        "search_text": "CursorPosition",
        "selection_mode": "start",
    },
    # Test 3: end 模式
    {
        "search_text": "EndPosition",
        "selection_mode": "end",
    },
    # Test 4: 模式切换 (使用自定义执行器)
    {
        "search_text": "ModeSwitch",
        "custom_executor": True,
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="select 模式 (高亮选区)",
        fixture_name="simple.docx",
        description="搜索 'Selection Test' 使用 select 模式，文本应被高亮选中",
        validator=validate_select_mode,
        tags=["selection_mode", "select"],
    ),
    TestCase(
        name="start 模式 (光标定位到开头)",
        fixture_name="simple.docx",
        description="搜索 'CursorPosition' 使用 start 模式，光标定位到匹配文本开头",
        validator=validate_start_mode,
        tags=["selection_mode", "start"],
    ),
    TestCase(
        name="end 模式 (光标定位到结尾)",
        fixture_name="simple.docx",
        description="搜索 'EndPosition' 使用 end 模式，光标定位到匹配文本结尾",
        validator=validate_end_mode,
        tags=["selection_mode", "end"],
    ),
    TestCase(
        name="模式切换验证",
        fixture_name="simple.docx",
        description="对 'ModeSwitch' 依次执行 select/start/end 三种模式，验证都能成功",
        validator=validate_mode_switching,
        tags=["selection_mode", "switching"],
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
            # Test 4 使用自定义执行器: 连续执行三种模式
            if config.get("custom_executor"):
                return await _run_mode_switching_test(
                    workspace, fixture, test_case, test_number, config
                )

            search_text = config["search_text"]
            selection_mode = config.get("selection_mode", "select")
            select_index = config.get("select_index", 1)
            search_options = config.get("search_options")

            print(f"\n📝 执行: 选择文本...")
            print(f"   搜索: '{search_text}'")
            print(f"   模式: {selection_mode}, 索引: {select_index}")

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


async def _run_mode_switching_test(
    workspace: Any,
    fixture: Any,
    test_case: TestCase,
    test_number: int,
    config: dict[str, Any],
) -> bool:
    """
    模式切换测试: 对同一文本依次执行 select/start/end 三种模式

    Args:
        workspace: Workspace 实例
        fixture: 文档夹具
        test_case: 测试用例
        test_number: 测试编号
        config: 选择配置

    Returns:
        是否通过
    """
    search_text = config["search_text"]
    modes = ["select", "start", "end"]
    mode_results: dict[str, dict[str, Any]] = {}

    for mode in modes:
        print(f"\n--- 执行 {mode} 模式 ---")
        print(f"   搜索: '{search_text}', 模式: {mode}")

        start_time = time.time()

        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": fixture.document_uri,
                "search_text": search_text,
                "selection_mode": mode,
                "select_index": 1,
            },
        )

        result = await workspace.execute(action)
        elapsed_ms = (time.time() - start_time) * 1000

        print(f"   ⏱️  执行时间: {elapsed_ms:.1f}ms")

        if result.success:
            data = result.data or {}
            print(f"   ✅ 成功, matchCount={data.get('matchCount', 0)}")
            mode_results[mode] = {
                "success": True,
                "matchCount": data.get("matchCount", 0),
                "data": data,
            }
        else:
            print(f"   ❌ 失败: {result.error}")
            mode_results[mode] = {
                "success": False,
                "matchCount": 0,
                "error": result.error,
            }

        # 模式之间等待一下
        await asyncio.sleep(1.0)

    # 将三次结果合并传递给 validator
    combined_data: dict[str, Any] = {"_mode_results": mode_results}

    print("\n📊 验证结果:")
    passed = True

    if test_case.validator:
        reader = DocumentReader(fixture.working_path)
        if not _call_validator(test_case.validator, combined_data, reader):
            passed = False

    print("\n" + "=" * 70)
    if passed:
        print(f"✅ 测试 {test_number} 通过")
    else:
        print(f"❌ 测试 {test_number} 失败")
    print("=" * 70)
    return passed


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
        description="Selection Modes Select Text E2E Tests (自动化版本)",
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
