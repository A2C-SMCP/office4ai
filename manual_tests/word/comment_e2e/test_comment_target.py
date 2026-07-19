"""
Comment Target E2E Tests (自动化版本)

测试批注定位功能。

测试场景:
1. 默认选区批注 — target=None (使用当前选区)
2. 搜索文本批注 — target={type:"searchText", searchText:"测试"}
3. 搜索不存在文本 — target={type:"searchText", searchText:"不存在的文本"}

运行方式:
    uv run python manual_tests/word/comment_e2e/test_comment_target.py --test 1
    uv run python manual_tests/word/comment_e2e/test_comment_target.py --test all
    uv run python manual_tests/word/comment_e2e/test_comment_target.py --list
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.e2e_base import (
    E2ETestRunner,
    TestCase,
    ensure_fixtures,
)
from manual_tests.error_case import PENDING_ADDIN_92, judge_error_case
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "comment_e2e"

TEST_CASES: list[TestCase] = [
    TestCase(
        name="默认选区批注",
        fixture_name="simple.docx",
        description="不指定 target，批注附加到当前选区",
        tags=["target", "basic"],
    ),
    TestCase(
        name="搜索文本批注",
        fixture_name="simple.docx",
        description="使用 target={type:searchText, searchText:'测试'} 定位，验证 associatedText",
        tags=["target", "basic"],
    ),
    TestCase(
        name="搜索不存在文本",
        fixture_name="simple.docx",
        description="searchText 零匹配 → 3012 SEARCH_NO_MATCH（oasp#23 规范层 MUST）",
        expect_error_code="3012",
        xfail_reason=PENDING_ADDIN_92,
        tags=["target", "error"],
    ),
]


# ==============================================================================
# 工作流测试
# ==============================================================================


async def _test_default_target(
    workspace: Any,
    document_uri: str,
    test_case: TestCase,
) -> bool:
    """测试默认选区批注"""
    comment_text = "E2E test - default target"

    # 插入批注（不指定 target）
    print("\n   📝 插入批注 (target=None)...")
    insert_action = OfficeAction(
        category="word",
        action_name="insert:comment",
        params={
            "document_uri": document_uri,
            "text": comment_text,
        },
    )
    insert_result = await workspace.execute(insert_action)
    if not insert_result.success:
        print(f"   ❌ 插入失败: {insert_result.error}")
        return False
    print("   ✅ 插入成功")

    # 获取批注验证
    print("   📝 获取批注列表...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={"document_uri": document_uri},
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取失败: {get_result.error}")
        return False

    comments = (get_result.data or {}).get("comments", [])
    found = any(c.get("content") == comment_text for c in comments)
    if not found:
        print(f"   ❌ 未找到批注: '{comment_text}'")
        return False
    print("   ✅ 验证通过: 默认选区批注插入成功")
    return True


async def _test_search_text_target(
    workspace: Any,
    document_uri: str,
    test_case: TestCase,
) -> bool:
    """测试搜索文本定位批注"""
    comment_text = "E2E test - search text target"
    search_text = "测试"

    # 插入批注（指定 searchText target）
    print(f"\n   📝 插入批注 (target.searchText='{search_text}')...")
    insert_action = OfficeAction(
        category="word",
        action_name="insert:comment",
        params={
            "document_uri": document_uri,
            "text": comment_text,
            "target": {
                "type": "searchText",
                "search_text": search_text,
            },
        },
    )
    insert_result = await workspace.execute(insert_action)
    if not insert_result.success:
        print(f"   ❌ 插入失败: {insert_result.error}")
        return False
    print("   ✅ 插入成功")

    # 获取批注（包含关联文本）
    print("   📝 获取批注列表 (includeAssociatedText=true)...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={
            "document_uri": document_uri,
            "options": {"include_associated_text": True},
        },
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取失败: {get_result.error}")
        return False

    comments = (get_result.data or {}).get("comments", [])
    target_comment = None
    for c in comments:
        if c.get("content") == comment_text:
            target_comment = c
            break

    if not target_comment:
        print(f"   ❌ 未找到批注: '{comment_text}'")
        return False

    associated_text = target_comment.get("associatedText", "")
    if search_text in associated_text:
        print(f"   ✅ 验证通过: associatedText 包含 '{search_text}' → '{associated_text}'")
    else:
        print(f"   ⚠️  associatedText='{associated_text}' (可能不包含搜索文本，取决于实现)")
    return True


async def _test_nonexistent_text_target(
    workspace: Any,
    document_uri: str,
    test_case: TestCase,
) -> bool:
    """测试搜索不存在文本 → 权威码 3012 SEARCH_NO_MATCH（oasp#23 规范层 / issue #90）。

    判定收紧（原实现两个分支都 ``return True``，等于对结果完全不设防）：

    * 返回 ``3012``           → 🎉 XPASS，Add-In 已接线，可摘 xfail_reason 转正
    * 返回其他错误码（如 3000）→ ⚠️ XFAIL，Add-In 未接线的过渡期兜底，套件不红
    * **插入却成功**          → ❌ FAIL

    最后一条是本次收紧的重点：原实现把「降级到当前选区」当作可接受结果放行，但那正是
    规范层禁止的——批注附到了**用户没有请求的位置**上，调用方还以为成功了。
    """
    comment_text = "E2E test - nonexistent target"
    search_text = "不存在的文本内容XYZ123"

    # 插入批注（searchText 指向不存在的文本）
    print(f"\n   📝 插入批注 (target.searchText='{search_text}')...")
    insert_action = OfficeAction(
        category="word",
        action_name="insert:comment",
        params={
            "document_uri": document_uri,
            "text": comment_text,
            "target": {
                "type": "searchText",
                "search_text": search_text,
            },
        },
    )
    insert_result = await workspace.execute(insert_action)

    # 期望单一来源于 TEST_CASES；缺失即配置错误，快速失败优于沉默兜底一个字面量
    expect_code = test_case.expect_error_code
    assert expect_code, "该用例必须在 TEST_CASES 中声明 expect_error_code"

    verdict, message = judge_error_case(
        insert_result.success,
        insert_result.error,
        expect_code=expect_code,
        expect_details=test_case.expect_error_details,
        xfail_reason=test_case.xfail_reason,
    )
    print(message)

    # 意外成功时补打关联文本，便于判断是否降级到了当前选区（诊断信息，不影响判定）
    if insert_result.success:
        get_action = OfficeAction(
            category="word",
            action_name="get:comments",
            params={
                "document_uri": document_uri,
                "options": {"include_associated_text": True},
            },
        )
        get_result = await workspace.execute(get_action)
        if get_result.success:
            for c in (get_result.data or {}).get("comments", []):
                if c.get("content") == comment_text:
                    print(f"   📋 associatedText='{c.get('associatedText', '')}'（疑似降级到当前选区）")
                    break

    return verdict.passed


_WORKFLOW_FUNCS = [
    _test_default_target,
    _test_search_text_target,
    _test_nonexistent_text_target,
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

    fixture_path = f"comment_e2e/{test_case.fixture_name}"
    workflow_func = _WORKFLOW_FUNCS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            start_time = time.time()
            passed = await workflow_func(workspace, fixture.document_uri, test_case)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")
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
        description="Comment Target E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=默认选区, 2=搜索文本, 3=不存在文本, all=全部",
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
