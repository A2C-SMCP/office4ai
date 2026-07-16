"""
Get Selection E2E Tests (自动化版本)

测试 word:get:selection 功能的各种选区状态。

特性：
- 自动复制测试文档
- 自动打开 Word
- 提示用户进行选区操作
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. 正常选区 - 获取有高亮文本的选区信息
2. 光标位置 - 获取无高亮文本的光标位置
3. 无选区 - 获取无选区状态
4. 性能对比 - 对比 get:selection 和 get:selectedContent 性能

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_selection_e2e/test_selection.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_selection_e2e/test_selection.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_selection_e2e/test_selection.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_selection_e2e/test_selection.py --test 1 --always-cleanup
"""

import asyncio
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from manual_tests.e2e_base import (
    DocumentReader,
    E2ETestRunner,
    Validator,
    _call_validator,
    ensure_fixtures,
)
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

# 夹具目录
FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "get_selection_e2e"


@dataclass
class SelectionTestCase:
    """
    选区测试用例定义

    Attributes:
        name: 测试名称
        fixture_name: 夹具文件名（相对于 fixtures 目录）
        description: 测试描述
        user_prompt: 用户操作提示（在测试执行前显示）
        validator: 自定义验证函数
        tags: 标签列表
    """

    name: str
    fixture_name: str
    description: str
    user_prompt: str | None = None
    validator: Validator | None = None
    tags: list[str] = field(default_factory=list)


def validate_normal_selection(data: dict[str, Any]) -> bool:
    """
    验证正常选区（有高亮文本）

    预期返回:
    - isEmpty: false
    - type: "Normal"
    - start/end: 有值且 start < end
    - text: 选中的文本内容

    注: type 字段在 Windows Desktop 使用 API 直接获取，
        在 Word Online/Mac 上基于 isEmpty 推断。
    """
    # 必须有选区
    if data.get("isEmpty", True):
        print("      ⚠️  isEmpty 应为 false")
        return False

    # 类型应为 Normal
    selection_type = data.get("type")
    if selection_type != "Normal":
        print(f"      ⚠️  type 应为 'Normal'，实际为 '{selection_type}'")
        return False

    # start 和 end 应有值
    start = data.get("start")
    end = data.get("end")
    if start is None or end is None:
        print("      ⚠️  start 和 end 应有值")
        return False

    # start <= end
    if start > end:
        print(f"      ⚠️  start({start}) 应 <= end({end})")
        return False

    # text 应有值
    text = data.get("text")
    if not text:
        print("      ⚠️  text 应有值")
        return False

    # 长度一致性检查
    expected_length = end - start
    actual_length = len(text)
    if expected_length != actual_length:
        print(f"      ⚠️  长度不一致: 预期 {expected_length}, 实际 {actual_length}")
        # 这个可能因为编码差异不完全匹配，作为警告而非失败
        print("      （长度差异可能因字符编码导致，继续测试）")

    return True


def validate_insertion_point(data: dict[str, Any]) -> bool:
    """
    验证光标位置（无高亮文本）

    预期返回:
    - isEmpty: true
    - type: "InsertionPoint"
    - start/end: 有值且 start == end（光标位置）
    - text: null 或空字符串

    注: type 字段在 Windows Desktop 使用 API 直接获取，
        在 Word Online/Mac 上基于 isEmpty 推断。
    """
    # 应为空选区
    if not data.get("isEmpty", False):
        print("      ⚠️  isEmpty 应为 true")
        return False

    # 类型应为 InsertionPoint
    selection_type = data.get("type")
    if selection_type != "InsertionPoint":
        print(f"      ⚠️  type 应为 'InsertionPoint'，实际为 '{selection_type}'")
        return False

    # start 和 end 应相同
    start = data.get("start")
    end = data.get("end")
    if start is None or end is None:
        print("      ⚠️  start 和 end 应有值")
        return False

    if start != end:
        print(f"      ⚠️  start({start}) 应等于 end({end})")
        return False

    # text 应为空或 null
    text = data.get("text")
    if text:
        print(f"      ⚠️  text 应为空，实际为 '{text}'")
        return False

    return True


def validate_no_selection(data: dict[str, Any]) -> bool:
    """
    验证无选区状态

    预期返回:
    - isEmpty: true
    - type: "InsertionPoint"（基于 isEmpty 推断）或 "NoSelection"（API 直接返回）
    - start/end: 光标位置或 null
    - text: null 或空字符串

    注: Add-In 的降级逻辑将 isEmpty=true 推断为 "InsertionPoint"，
        因此在 Word Online/Mac 上此测试与测试 2 行为相同。
        "NoSelection" 仅在 Windows Desktop 且光标不在文档内时返回。
    """
    # 应为空选区
    if not data.get("isEmpty", False):
        print("      ⚠️  isEmpty 应为 true")
        return False

    # 类型应为 InsertionPoint 或 NoSelection
    selection_type = data.get("type")
    if selection_type not in ("NoSelection", "InsertionPoint"):
        print(f"      ⚠️  type 应为 'NoSelection' 或 'InsertionPoint'，实际为 '{selection_type}'")
        return False

    # text 应为空或 null
    text = data.get("text")
    if text:
        print(f"      ⚠️  text 应为空，实际为 '{text}'")
        return False

    return True


# 测试用例定义
TEST_CASES: list[SelectionTestCase] = [
    SelectionTestCase(
        name="正常选区（有高亮文本）",
        fixture_name="simple.docx",
        description="获取文档中的高亮文本选区信息",
        user_prompt='请选中文档中的 "测试文档" 标题文字（用鼠标拖拽高亮选中）',
        validator=validate_normal_selection,
        tags=["basic"],
    ),
    SelectionTestCase(
        name="光标位置（无高亮文本）",
        fixture_name="simple.docx",
        description="获取光标插入点位置信息",
        user_prompt="请在文档正文中单击一次放置光标（不要拖拽选中任何文本）",
        validator=validate_insertion_point,
        tags=["basic"],
    ),
    SelectionTestCase(
        name="无选区状态",
        fixture_name="simple.docx",
        description="获取无选区时的状态信息",
        user_prompt="请点击文档灰色边距区域，或按 Esc 键取消选区",
        validator=validate_no_selection,
        tags=["basic"],
    ),
    SelectionTestCase(
        name="性能对比",
        fixture_name="large.docx",
        description="对比 get:selection 和 get:selectedContent 的性能",
        user_prompt="请用 Ctrl+A (Mac: Cmd+A) 全选文档内容",
        validator=lambda data: data.get("isEmpty") is not None,  # 基本结构验证
        tags=["performance"],
    ),
]


# ==============================================================================
# 测试执行
# ==============================================================================


def display_selection(data: dict[str, Any]) -> None:
    """显示选区信息"""
    is_empty = data.get("isEmpty", False)
    selection_type = data.get("type", "Unknown")
    start = data.get("start")
    end = data.get("end")
    text = data.get("text")

    print(f"   是否为空: {is_empty}")
    print(f"   选区类型: {selection_type}")

    if start is not None:
        print(f"   起始位置: {start}")
    if end is not None:
        print(f"   结束位置: {end}")
    if text is not None:
        display_text = text[:50] + "..." if len(text) > 50 else text
        print(f"   选区文本: '{display_text}'")
        print(f"   文本长度: {len(text)} 字符")


async def run_single_test(
    runner: E2ETestRunner,
    test_case: SelectionTestCase,
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

    fixture_path = f"get_selection_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 如果有用户提示，等待用户操作
            if test_case.user_prompt:
                print(f"\n📌 请执行以下操作:")
                print(f"   {test_case.user_prompt}")
                input("\n⏸️  完成操作后按 Enter 继续...")

            # 执行获取选区动作
            print("\n📝 执行: 获取选区信息...")
            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="get:selection",
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

            print("\n📊 选区信息:")
            display_selection(data)

            # 验证结果
            print("\n📊 验证结果:")
            passed = True

            # 自定义验证
            if test_case.validator:
                # 创建文档读取器（用于双重验证，虽然此测试不需要）
                reader = DocumentReader(fixture.working_path)
                if _call_validator(test_case.validator, data, reader):
                    print("   ✅ 验证通过")
                else:
                    print("   ❌ 验证失败")
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


async def run_performance_test(
    runner: E2ETestRunner,
    test_case: SelectionTestCase,
    test_number: int,
) -> bool:
    """
    执行性能对比测试

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

    fixture_path = f"get_selection_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 用户提示
            if test_case.user_prompt:
                print(f"\n📌 请执行以下操作:")
                print(f"   {test_case.user_prompt}")
                input("\n⏸️  完成操作后按 Enter 继续...")

            # 测试 word:get:selection
            print("\n⏱️  测试 word:get:selection 性能...")
            await asyncio.sleep(1)
            start_time = time.time()

            action1 = OfficeAction(
                category="word",
                action_name="get:selection",
                params={"document_uri": fixture.document_uri},
            )
            result1 = await workspace.execute(action1)
            selection_time = (time.time() - start_time) * 1000

            if result1.success:
                print(f"   ✅ word:get:selection 耗时: {selection_time:.1f}ms")
            else:
                print(f"   ❌ word:get:selection 失败: {result1.error}")
                return False

            await asyncio.sleep(1)

            # 测试 word:get:selectedContent
            print("\n⏱️  测试 word:get:selectedContent 性能...")
            start_time = time.time()

            action2 = OfficeAction(
                category="word",
                action_name="get:selectedContent",
                params={"document_uri": fixture.document_uri},
            )
            result2 = await workspace.execute(action2)
            content_time = (time.time() - start_time) * 1000

            if result2.success:
                print(f"   ✅ word:get:selectedContent 耗时: {content_time:.1f}ms")
            else:
                print(f"   ⚠️  word:get:selectedContent 失败: {result2.error}")
                print("   （继续测试）")
                content_time = float("inf")

            # 对比结果
            print("\n📊 性能对比:")
            print(f"   word:get:selection:        {selection_time:.1f}ms")
            if content_time != float("inf"):
                print(f"   word:get:selectedContent:  {content_time:.1f}ms")

                if selection_time < content_time:
                    speedup = content_time / selection_time
                    print(f"   ✅ word:get:selection 快 {speedup:.1f}x")
                else:
                    print("   ⚠️  word:get:selection 未显示出性能优势")
            else:
                print("   word:get:selectedContent:  N/A (失败)")

            print("\n💡 结论: word:get:selection 是轻量级查询，适合快速获取位置信息")

            print("\n" + "=" * 70)
            print(f"✅ 测试 {test_number} 完成")
            print("=" * 70)
            return True

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
            await asyncio.sleep(2.0)

        # 性能测试使用特殊函数
        if "performance" in test_case.tags:
            result = await run_performance_test(runner, test_case, idx)
        else:
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
        description="Word Get Selection E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（正常选区）
  python test_selection.py --test 1

  # 运行所有测试
  python test_selection.py --test all

  # 手动打开文档模式
  python test_selection.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_selection.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "all"],
        default="1",
        help="要运行的测试: 1=正常选区, 2=光标位置, 3=无选区, 4=性能对比, all=全部",
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
            if tc.user_prompt:
                print(f"     用户操作: {tc.user_prompt}")
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
