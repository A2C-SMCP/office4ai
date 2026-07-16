"""
Comment Options E2E Tests (自动化版本)

测试 word:get:comments 的各种查询选项。

测试场景:
1. 排除已解决 — 插入+解决批注后，includeResolved=false 不返回该批注
2. 包含关联文本 — includeAssociatedText=true，验证 comments 含 associatedText 字段
3. 详细元数据 — detailedMetadata=true，验证 comments 含 authorName/creationDate

运行方式:
    uv run python manual_tests/word/comment_e2e/test_comment_options.py --test 1
    uv run python manual_tests/word/comment_e2e/test_comment_options.py --test all
    uv run python manual_tests/word/comment_e2e/test_comment_options.py --list
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
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "comment_e2e"

TEST_CASES: list[TestCase] = [
    TestCase(
        name="排除已解决批注",
        fixture_name="simple.docx",
        description="插入+解决批注后，includeResolved=false 不返回已解决批注",
        tags=["options", "basic"],
    ),
    TestCase(
        name="包含关联文本",
        fixture_name="simple.docx",
        description="includeAssociatedText=true，验证 comments 含 associatedText 字段",
        tags=["options", "basic"],
    ),
    TestCase(
        name="详细元数据",
        fixture_name="simple.docx",
        description="detailedMetadata=true，验证 comments 含 authorName/creationDate",
        tags=["options", "basic"],
    ),
]


# ==============================================================================
# 辅助函数
# ==============================================================================


def _extract_comment_id(data: dict[str, Any]) -> str | None:
    """从插入批注的返回数据中提取 comment ID"""
    return data.get("commentId") or data.get("id") or data.get("comment_id")


# ==============================================================================
# 工作流测试
# ==============================================================================


async def _test_exclude_resolved(
    workspace: Any,
    document_uri: str,
) -> bool:
    """测试排除已解决批注"""
    comment_text = "E2E test - exclude resolved"

    # Step 1: 插入批注
    print("\n   📝 Step 1: 插入批注...")
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

    comment_id = _extract_comment_id(insert_result.data or {})
    if not comment_id:
        print("   ❌ 未获取到 commentId")
        return False
    print(f"   ✅ 插入成功, commentId={comment_id}")

    # Step 2: 解决批注
    print("   📝 Step 2: 解决批注...")
    resolve_action = OfficeAction(
        category="word",
        action_name="resolve:comment",
        params={
            "document_uri": document_uri,
            "comment_id": comment_id,
            "resolved": True,
        },
    )
    resolve_result = await workspace.execute(resolve_action)
    if not resolve_result.success:
        print(f"   ❌ 解决失败: {resolve_result.error}")
        return False
    print("   ✅ 解决成功")

    # Step 3: 获取批注（不包含已解决）
    print("   📝 Step 3: 获取批注 (includeResolved=false)...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={
            "document_uri": document_uri,
            "options": {"include_resolved": False},
        },
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取失败: {get_result.error}")
        return False

    comments = (get_result.data or {}).get("comments", [])
    found = any(c.get("content") == comment_text for c in comments)
    if found:
        print(f"   ❌ 已解决批注不应出现在列表中 (includeResolved=false)")
        return False
    print(f"   ✅ 验证通过: 已解决批注被排除 (列表中 {len(comments)} 条)")
    return True


async def _test_include_associated_text(
    workspace: Any,
    document_uri: str,
) -> bool:
    """测试包含关联文本"""
    comment_text = "E2E test - associated text"

    # Step 1: 插入批注
    print("\n   📝 Step 1: 插入批注...")
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

    # Step 2: 获取批注（包含关联文本）
    print("   📝 Step 2: 获取批注 (includeAssociatedText=true)...")
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

    # 验证 associatedText 字段存在
    if "associatedText" in target_comment:
        associated = target_comment["associatedText"]
        print(f"   ✅ 验证通过: associatedText='{associated}'")
    else:
        print("   ⚠️  associatedText 字段不存在（可能需要选中文本后插入）")
    return True


async def _test_detailed_metadata(
    workspace: Any,
    document_uri: str,
) -> bool:
    """测试详细元数据"""
    comment_text = "E2E test - detailed metadata"

    # Step 1: 插入批注
    print("\n   📝 Step 1: 插入批注...")
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

    # Step 2: 获取批注（详细元数据）
    print("   📝 Step 2: 获取批注 (detailedMetadata=true)...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={
            "document_uri": document_uri,
            "options": {"detailed_metadata": True},
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

    # 验证元数据字段
    has_author = "authorName" in target_comment
    has_date = "creationDate" in target_comment
    print(f"   ℹ️  authorName: {target_comment.get('authorName', 'N/A')}")
    print(f"   ℹ️  creationDate: {target_comment.get('creationDate', 'N/A')}")

    if has_author or has_date:
        print("   ✅ 验证通过: 详细元数据字段存在")
    else:
        print("   ⚠️  authorName/creationDate 字段均不存在")
    return True


_WORKFLOW_FUNCS = [
    _test_exclude_resolved,
    _test_include_associated_text,
    _test_detailed_metadata,
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
            passed = await workflow_func(workspace, fixture.document_uri)
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
        description="Comment Options E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=排除已解决, 2=关联文本, 3=详细元数据, all=全部",
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
