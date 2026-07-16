"""
Location-based Insert Text E2E Tests (自动化版本)

测试不同插入位置（location 参数）的文本插入功能。

测试场景:
1. location="Start" - 在文档开头插入
2. location="End" - 在文档末尾插入
3. location="Cursor" - 在光标位置插入（需用户定位光标）
4. 连续多次插入测试位置累积效果

运行方式:
    uv run python manual_tests/insert_text_e2e/test_location_insert.py --test 1
    uv run python manual_tests/insert_text_e2e/test_location_insert.py --test all
    uv run python manual_tests/insert_text_e2e/test_location_insert.py --list
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


def validate_start_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证文档开头插入"""
    reader.reload()
    if not reader.paragraph_starts_with(0, "[文档开头标题]"):
        print("   ❌ 第一段不以 '[文档开头标题]' 开头")
        print(f"   实际第一段: '{reader.get_paragraph(0)}'")
        return False
    print("   ✅ 文档内容验证通过: 第一段以插入文本开头")
    return True


def validate_end_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证文档末尾插入"""
    reader.reload()
    if not reader.contains("[文档末尾追加内容]"):
        print("   ❌ 文档中未找到 '[文档末尾追加内容]'")
        return False
    # 检查是否在最后的段落中
    paragraphs = reader.paragraphs
    found_at_end = any("[文档末尾追加内容]" in p for p in paragraphs[-3:])
    if not found_at_end:
        print("   ❌ '[文档末尾追加内容]' 不在文档末尾区域")
        return False
    print("   ✅ 文档内容验证通过: 末尾包含插入文本")
    return True


def validate_cursor_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证光标位置插入"""
    reader.reload()
    if not reader.contains("[光标位置文本]"):
        print("   ❌ 文档中未找到 '[光标位置文本]'")
        return False
    print("   ✅ 文档内容验证通过: 包含光标位置插入的文本")
    return True


def validate_multiple_inserts(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证连续多次插入"""
    reader.reload()
    texts = ["第一次插入（开头）", "第二次插入（末尾）", "第三次插入（末尾）"]
    for text in texts:
        if not reader.contains(text):
            print(f"   ❌ 文档中未找到 '{text}'")
            return False
    print("   ✅ 文档内容验证通过: 包含所有三次插入的文本")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[TestCase] = [
    TestCase(
        name="在文档开头插入文本",
        fixture_name="simple.docx",
        description="使用 location='Start' 在已有内容的文档开头插入文本，验证第一段以插入文本开头",
        validator=validate_start_insert,
        tags=["location"],
    ),
    TestCase(
        name="在文档末尾插入文本",
        fixture_name="simple.docx",
        description="使用 location='End' 在文档末尾插入文本，验证末尾包含插入文本",
        validator=validate_end_insert,
        tags=["location"],
    ),
    TestCase(
        name="在光标位置插入文本",
        fixture_name="simple.docx",
        description="使用 location='Cursor' 在用户定位的光标处插入文本（需用户交互）",
        validator=validate_cursor_insert,
        tags=["location", "interactive"],
    ),
    TestCase(
        name="连续多次插入",
        fixture_name="simple.docx",
        description="依次在 Start/End/End 位置插入 3 次，验证所有文本都存在",
        validator=validate_multiple_inserts,
        tags=["location", "multi"],
    ),
]

# 每个测试用例的执行逻辑不同，使用 dispatch 模式


# ==============================================================================
# 测试执行
# ==============================================================================


async def _execute_test_1(workspace: Any, fixture: Any) -> tuple[bool, dict[str, Any]]:
    """测试 1: Start 位置插入"""
    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params={
            "document_uri": fixture.document_uri,
            "text": "[文档开头标题]\n",
            "location": "Start",
        },
    )
    result = await workspace.execute(action)
    return result.success, result.data or {}


async def _execute_test_2(workspace: Any, fixture: Any) -> tuple[bool, dict[str, Any]]:
    """测试 2: End 位置插入"""
    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params={
            "document_uri": fixture.document_uri,
            "text": "\n[文档末尾追加内容]",
            "location": "End",
        },
    )
    result = await workspace.execute(action)
    return result.success, result.data or {}


async def _execute_test_3(workspace: Any, fixture: Any) -> tuple[bool, dict[str, Any]]:
    """测试 3: Cursor 位置插入（需用户交互）"""
    print("\n   ⚠️  请将光标移动到文档中的任意位置")
    input("   按 Enter 继续...")

    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params={
            "document_uri": fixture.document_uri,
            "text": "[光标位置文本]",
            "location": "Cursor",
        },
    )
    result = await workspace.execute(action)
    return result.success, result.data or {}


async def _execute_test_4(workspace: Any, fixture: Any) -> tuple[bool, dict[str, Any]]:
    """测试 4: 连续多次插入"""
    inserts = [
        ("Start", "=== 第一次插入（开头） ===\n"),
        ("End", "\n=== 第二次插入（末尾） ==="),
        ("End", "\n=== 第三次插入（末尾） ==="),
    ]

    print(f"\n   将执行 {len(inserts)} 次连续插入")
    for i, (location, text) in enumerate(inserts, 1):
        print(f"   {i}. location={location}: {text.strip()}")

    last_data: dict[str, Any] = {}
    for i, (location, text) in enumerate(inserts, 1):
        print(f"\n   --- 执行第 {i} 次插入 ---")
        action = OfficeAction(
            category="word",
            action_name="insert:text",
            params={
                "document_uri": fixture.document_uri,
                "text": text,
                "location": location,
            },
        )
        result = await workspace.execute(action)
        if not result.success:
            print(f"   ❌ 第 {i} 次插入失败: {result.error}")
            return False, {}
        print(f"   ✅ 第 {i} 次插入成功")
        last_data = result.data or {}
        if i < len(inserts):
            await asyncio.sleep(1)

    return True, last_data


_EXECUTORS = [_execute_test_1, _execute_test_2, _execute_test_3, _execute_test_4]


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

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            print(f"\n📝 执行: {test_case.name}...")
            start_time = time.time()

            executor = _EXECUTORS[test_number - 1]
            success, data = await executor(workspace, fixture)

            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 操作失败")
                return False

            print("✅ 协议返回成功")

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
        description="Location-based Insert Text E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=Start, 2=End, 3=Cursor, 4=多次插入, all=全部",
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
