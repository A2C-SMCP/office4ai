"""
Edge Cases Replace Selection E2E Tests (自动化版本)

测试 word:replace:selection 的边界情况和错误处理。

测试场景:
1. 空选区错误 - 不选中文本直接替换，预期返回错误码 3002
2. 替换为空字符串（删除选中内容）- 选中文本后替换为空文本
3. 替换为图片 - 选中文本后替换为 base64 编码的图片

运行方式:
    # 运行单个测试
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test 1

    # 运行所有测试
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/replace_selection_e2e/test_edge_cases.py --test 1 --always-cleanup
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


def validate_empty_string_replace(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证空字符串替换（删除选中内容）"""
    reader.reload()
    # 空字符串替换成功即视为通过，具体内容需人工确认选中的文本已被删除
    print("   ✅ 协议返回成功: 选中内容已被删除")
    print("   👀 请在 Word 中确认选中的文本已被删除")
    return True


def validate_image_replace(data: dict[str, Any]) -> bool:
    """验证图片替换（DataValidator，python-docx 无法验证 Add-In 插入的图片）"""
    # 协议返回成功即视为通过
    print("   ✅ 协议返回成功: 图片替换完成")
    print("   👀 请在 Word 中确认图片已替换选中内容")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[TestCase] = [
    TestCase(
        name="空选区错误",
        fixture_name="simple.docx",
        description="不选中任何文本直接替换，预期返回错误码 3002 (SELECTION_EMPTY)",
        # 无 validator: 此测试使用自定义执行器，直接检查 result.success == False
        tags=["edge", "error"],
    ),
    TestCase(
        name="替换为空字符串（删除选中内容）",
        fixture_name="simple.docx",
        description="选中文本后替换为空字符串，验证选中内容被删除",
        validator=validate_empty_string_replace,
        tags=["edge", "delete"],
    ),
    TestCase(
        name="替换为图片",
        fixture_name="simple.docx",
        description="选中文本后替换为 base64 编码的图片（python-docx 无法验证，仅检查协议返回）",
        validator=validate_image_replace,
        tags=["edge", "image"],
    ),
]

# 每个测试用例的替换内容
_REPLACE_CONTENTS: list[dict[str, Any]] = [
    # Test 1: 空选区（内容不重要，预期失败）
    {"text": "This should fail"},
    # Test 2: 空字符串
    {"text": ""},
    # Test 3: 图片
    {
        "images": [
            {
                "base64": (
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAA"
                    "DUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=="
                ),
                "width": 100,
                "height": 100,
                "altText": "Test Image",
            }
        ]
    },
]


# ==============================================================================
# 自定义执行器
# ==============================================================================


async def _execute_empty_selection_test(
    workspace: Any,
    fixture: Any,
) -> tuple[bool, str]:
    """
    测试 1 自定义执行器: 空选区错误

    不让用户选中文本，直接发送替换请求，验证返回错误码 3002。

    Returns:
        (passed, message): 是否通过和结果消息
    """
    print("\n   ⚠️  请确保 Word 中【没有】选中任何文本（点击空白处取消选择）")
    input("   确认没有选中文本后按 Enter...")

    action = OfficeAction(
        category="word",
        action_name="replace:selection",
        params={
            "document_uri": fixture.document_uri,
            "content": _REPLACE_CONTENTS[0],
        },
    )

    result = await workspace.execute(action)

    if not result.success:
        error_str = str(result.error or "")
        if "3002" in error_str:
            return True, f"✅ 正确返回错误码 3002: {result.error}"
        else:
            return False, f"⚠️  返回了错误但错误码不匹配 (预期 3002): {result.error}"
    else:
        return False, "❌ 预期返回错误，但请求成功了（可能 Word 中有选中的内容）"


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

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            start_time = time.time()

            # 测试 1: 空选区错误 — 使用自定义执行器
            if test_number == 1:
                print(f"\n📝 执行: {test_case.name}...")
                passed, message = await _execute_empty_selection_test(workspace, fixture)
                elapsed_ms = (time.time() - start_time) * 1000
                print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")
                print(f"\n📊 验证结果:")
                print(f"   {message}")

                print("\n" + "=" * 70)
                if passed:
                    print(f"✅ 测试 {test_number} 通过")
                else:
                    print(f"❌ 测试 {test_number} 失败")
                print("=" * 70)
                return passed

            # 测试 2, 3: 需要用户先选中文本
            input("\n请在 Word 中选中一些文本后按 Enter...")

            print(f"\n📝 执行: {test_case.name}...")
            content = _REPLACE_CONTENTS[test_number - 1]

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

            # Validator 验证
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
        description="Edge Cases Replace Selection E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（空选区错误）
  python test_edge_cases.py --test 1

  # 运行所有测试
  python test_edge_cases.py --test all

  # 手动打开文档模式
  python test_edge_cases.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_edge_cases.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=空选区, 2=空字符串, 3=图片, all=全部",
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
