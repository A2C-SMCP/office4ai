"""
Comment CRUD E2E Tests (自动化版本)

测试批注的核心 CRUD 工作流（有状态链式操作）。

测试场景:
1. 插入+获取 — insert:comment → get:comments, 验证批注内容
2. 插入+回复+获取 — insert:comment → reply:comment → get:comments, 验证回复
3. 插入+解决+获取 — insert:comment → resolve:comment → get:comments, 验证 resolved=True
4. 插入+删除+获取 — insert:comment → delete:comment → get:comments, 验证批注已删除

运行方式:
    uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test 1
    uv run python manual_tests/word/comment_e2e/test_comment_crud.py --test all
    uv run python manual_tests/word/comment_e2e/test_comment_crud.py --list
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
        name="插入+获取批注",
        fixture_name="simple.docx",
        description="插入一条批注后获取批注列表，验证批注内容存在",
        tags=["crud", "basic"],
    ),
    TestCase(
        name="插入+回复+获取批注",
        fixture_name="simple.docx",
        description="插入批注 → 回复批注 → 获取批注列表，验证回复内容存在",
        tags=["crud", "basic"],
    ),
    TestCase(
        name="插入+解决+获取批注",
        fixture_name="simple.docx",
        description="插入批注 → 解决批注 → 获取批注列表，验证 resolved=True",
        tags=["crud", "basic"],
    ),
    TestCase(
        name="插入+删除+获取批注",
        fixture_name="simple.docx",
        description="插入批注 → 删除批注 → 获取批注列表，验证批注已被删除",
        tags=["crud", "basic"],
    ),
]


# ==============================================================================
# 工作流测试
# ==============================================================================


def _extract_comment_id(data: dict[str, Any]) -> str | None:
    """从插入批注的返回数据中提取 comment ID"""
    return data.get("commentId") or data.get("id") or data.get("comment_id")


async def _test_insert_and_get(
    workspace: Any,
    document_uri: str,
) -> bool:
    """工作流 1: 插入 + 获取"""
    comment_text = "E2E test comment - insert and get"

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
        print(f"   ❌ 插入批注失败: {insert_result.error}")
        return False

    insert_data = insert_result.data or {}
    comment_id = _extract_comment_id(insert_data)
    print(f"   ✅ 插入成功, commentId={comment_id}")

    # Step 2: 获取批注
    print("   📝 Step 2: 获取批注列表...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={"document_uri": document_uri},
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取批注失败: {get_result.error}")
        return False

    get_data = get_result.data or {}
    comments = get_data.get("comments", [])
    print(f"   ✅ 获取成功, 共 {len(comments)} 条批注")

    # Step 3: 验证
    found = any(c.get("content") == comment_text for c in comments)
    if not found:
        print(f"   ❌ 未找到插入的批注文本: '{comment_text}'")
        print(f"   实际批注: {comments}")
        return False
    print(f"   ✅ 验证通过: 找到批注 '{comment_text}'")
    return True


async def _test_insert_reply_get(
    workspace: Any,
    document_uri: str,
) -> bool:
    """工作流 2: 插入 + 回复 + 获取"""
    comment_text = "E2E test comment - for reply"
    reply_text = "E2E test reply content"

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
        print(f"   ❌ 插入批注失败: {insert_result.error}")
        return False

    insert_data = insert_result.data or {}
    comment_id = _extract_comment_id(insert_data)
    if not comment_id:
        print("   ❌ 未获取到 commentId")
        return False
    print(f"   ✅ 插入成功, commentId={comment_id}")

    # Step 2: 回复批注
    print("   📝 Step 2: 回复批注...")
    reply_action = OfficeAction(
        category="word",
        action_name="reply:comment",
        params={
            "document_uri": document_uri,
            "comment_id": comment_id,
            "text": reply_text,
        },
    )
    reply_result = await workspace.execute(reply_action)
    if not reply_result.success:
        print(f"   ❌ 回复批注失败: {reply_result.error}")
        return False
    print("   ✅ 回复成功")

    # Step 3: 获取批注（包含回复）
    print("   📝 Step 3: 获取批注列表 (includeReplies=true)...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={
            "document_uri": document_uri,
            "options": {"include_replies": True},
        },
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取批注失败: {get_result.error}")
        return False

    get_data = get_result.data or {}
    comments = get_data.get("comments", [])

    # 查找目标批注
    target_comment = None
    for c in comments:
        if c.get("content") == comment_text:
            target_comment = c
            break

    if not target_comment:
        print(f"   ❌ 未找到原始批注: '{comment_text}'")
        return False

    # 验证回复
    replies = target_comment.get("replies", [])
    has_reply = any(r.get("content") == reply_text for r in replies)
    if not has_reply:
        print(f"   ❌ 未找到回复: '{reply_text}'")
        print(f"   回复列表: {replies}")
        return False
    print(f"   ✅ 验证通过: 找到回复 '{reply_text}'")
    return True


async def _test_insert_resolve_get(
    workspace: Any,
    document_uri: str,
) -> bool:
    """工作流 3: 插入 + 解决 + 获取"""
    comment_text = "E2E test comment - for resolve"

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
        print(f"   ❌ 插入批注失败: {insert_result.error}")
        return False

    insert_data = insert_result.data or {}
    comment_id = _extract_comment_id(insert_data)
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
        print(f"   ❌ 解决批注失败: {resolve_result.error}")
        return False
    print("   ✅ 解决成功")

    # Step 3: 获取批注（包含已解决）
    print("   📝 Step 3: 获取批注列表 (includeResolved=true)...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={
            "document_uri": document_uri,
            "options": {"include_resolved": True},
        },
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取批注失败: {get_result.error}")
        return False

    get_data = get_result.data or {}
    comments = get_data.get("comments", [])

    # 查找目标批注并验证 resolved 状态
    target_comment = None
    for c in comments:
        if c.get("content") == comment_text:
            target_comment = c
            break

    if not target_comment:
        print(f"   ❌ 未找到批注: '{comment_text}'")
        return False

    if not target_comment.get("resolved"):
        print(f"   ❌ 批注未标记为已解决: resolved={target_comment.get('resolved')}")
        return False
    print("   ✅ 验证通过: 批注已标记为 resolved=True")
    return True


async def _test_insert_delete_get(
    workspace: Any,
    document_uri: str,
) -> bool:
    """工作流 4: 插入 + 删除 + 获取"""
    comment_text = "E2E test comment - for delete"

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
        print(f"   ❌ 插入批注失败: {insert_result.error}")
        return False

    insert_data = insert_result.data or {}
    comment_id = _extract_comment_id(insert_data)
    if not comment_id:
        print("   ❌ 未获取到 commentId")
        return False
    print(f"   ✅ 插入成功, commentId={comment_id}")

    # Step 2: 删除批注
    print("   📝 Step 2: 删除批注...")
    delete_action = OfficeAction(
        category="word",
        action_name="delete:comment",
        params={
            "document_uri": document_uri,
            "comment_id": comment_id,
        },
    )
    delete_result = await workspace.execute(delete_action)
    if not delete_result.success:
        print(f"   ❌ 删除批注失败: {delete_result.error}")
        return False
    print("   ✅ 删除成功")

    # Step 3: 获取批注
    print("   📝 Step 3: 获取批注列表...")
    get_action = OfficeAction(
        category="word",
        action_name="get:comments",
        params={"document_uri": document_uri},
    )
    get_result = await workspace.execute(get_action)
    if not get_result.success:
        print(f"   ❌ 获取批注失败: {get_result.error}")
        return False

    get_data = get_result.data or {}
    comments = get_data.get("comments", [])

    # 验证批注已被删除
    found = any(c.get("content") == comment_text for c in comments)
    if found:
        print(f"   ❌ 批注未被删除，仍在列表中: '{comment_text}'")
        return False
    print(f"   ✅ 验证通过: 批注已被删除 (当前 {len(comments)} 条)")
    return True


# 工作流函数映射
_WORKFLOW_FUNCS = [
    _test_insert_and_get,
    _test_insert_reply_get,
    _test_insert_resolve_get,
    _test_insert_delete_get,
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
        description="Comment CRUD E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=插入获取, 2=插入回复, 3=插入解决, 4=插入删除, all=全部",
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
