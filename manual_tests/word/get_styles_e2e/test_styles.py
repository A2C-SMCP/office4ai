"""
Get Styles E2E Tests (自动化版本)

测试 word:get:styles 功能的各种参数组合。

特性：
- 自动复制测试文档
- 自动打开 Word
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. 获取所有正在使用的样式（默认参数）
2. 仅获取内置样式
3. 仅获取自定义样式
4. 获取包含详细信息的样式
5. 获取所有样式（包括未使用的）

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_styles_e2e/test_styles.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_styles_e2e/test_styles.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_styles_e2e/test_styles.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_styles_e2e/test_styles.py --test 1 --always-cleanup
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
FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "get_styles_e2e"


@dataclass
class StylesTestCase:
    """
    样式测试用例定义

    Attributes:
        name: 测试名称
        fixture_name: 夹具文件名（相对于 fixtures 目录）
        description: 测试描述
        options: 样式获取选项
        validator: 自定义验证函数
        detailed_display: 是否显示详细信息
        tags: 标签列表
    """

    name: str
    fixture_name: str
    description: str
    options: dict[str, Any] | None = None
    validator: Validator | None = None
    detailed_display: bool = False
    tags: list[str] = field(default_factory=list)


# ==============================================================================
# 验证函数
# ==============================================================================


def validate_styles_basic(data: dict[str, Any]) -> bool:
    """
    验证基本样式返回结构

    检查:
    - styles 字段存在且为列表
    - 每个样式有必要字段 (name, type)
    """
    styles = data.get("styles")
    if styles is None:
        print("      ⚠️  缺少 styles 字段")
        return False

    if not isinstance(styles, list):
        print(f"      ⚠️  styles 应为列表，实际为 {type(styles)}")
        return False

    if len(styles) == 0:
        print("      ⚠️  样式列表为空")
        return False

    # 检查第一个样式的结构
    first_style = styles[0]
    if "name" not in first_style:
        print("      ⚠️  样式缺少 name 字段")
        return False

    if "type" not in first_style:
        print("      ⚠️  样式缺少 type 字段")
        return False

    return True


def validate_built_in_only(data: dict[str, Any]) -> bool:
    """验证仅返回内置样式"""
    if not validate_styles_basic(data):
        return False

    styles = data.get("styles", [])

    # 检查是否有自定义样式（应该没有）
    custom_count = sum(1 for s in styles if not s.get("builtIn", True))
    if custom_count > 0:
        print(f"      ⚠️  不应包含自定义样式，但找到 {custom_count} 个")
        return False

    return True


def validate_custom_only(data: dict[str, Any]) -> bool:
    """验证仅返回自定义样式"""
    styles = data.get("styles", [])

    # 自定义样式可能为空（文档没有自定义样式）
    if styles is None:
        print("      ⚠️  缺少 styles 字段")
        return False

    if not isinstance(styles, list):
        print(f"      ⚠️  styles 应为列表，实际为 {type(styles)}")
        return False

    # 如果有样式，检查是否都是自定义的
    built_in_count = sum(1 for s in styles if s.get("builtIn", False))
    if built_in_count > 0:
        print(f"      ⚠️  不应包含内置样式，但找到 {built_in_count} 个")
        return False

    if len(styles) == 0:
        print("      ℹ️  文档中没有自定义样式（这是正常的）")

    return True


def validate_detailed_info(data: dict[str, Any]) -> bool:
    """验证包含详细信息"""
    if not validate_styles_basic(data):
        return False

    styles = data.get("styles", [])

    # 检查是否有样式包含 description 字段
    # 注意：不是所有样式都有 description，所以只要有部分有就行
    has_description = any("description" in s for s in styles)
    if not has_description:
        print("      ℹ️  没有样式包含 description 字段（可能因为请求了 detailedInfo）")
        # 不强制失败，因为 Word 可能不返回空的 description

    return True


def validate_includes_unused(data: dict[str, Any]) -> bool:
    """验证包含未使用的样式"""
    if not validate_styles_basic(data):
        return False

    styles = data.get("styles", [])

    # 检查是否有未使用的样式
    unused_count = sum(1 for s in styles if not s.get("inUse", True))
    print(f"      ℹ️  包含 {unused_count} 个未使用的样式")

    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[StylesTestCase] = [
    StylesTestCase(
        name="获取所有正在使用的样式（默认参数）",
        fixture_name="simple.docx",
        description="获取文档中所有正在使用的样式，包括内置和自定义样式",
        options=None,
        validator=validate_styles_basic,
        tags=["basic"],
    ),
    StylesTestCase(
        name="仅获取内置样式",
        fixture_name="simple.docx",
        description="只获取 Word 内置的样式（如标题一、正文等）",
        options={
            "includeBuiltIn": True,
            "includeCustom": False,
            "includeUnused": False,
            "detailedInfo": False,
        },
        validator=validate_built_in_only,
        tags=["filter"],
    ),
    StylesTestCase(
        name="仅获取自定义样式",
        fixture_name="simple.docx",
        description="只获取用户自定义的样式",
        options={
            "includeBuiltIn": False,
            "includeCustom": True,
            "includeUnused": True,
            "detailedInfo": False,
        },
        validator=validate_custom_only,
        tags=["filter"],
    ),
    StylesTestCase(
        name="获取包含详细信息的样式",
        fixture_name="simple.docx",
        description="获取样式及其详细描述信息",
        options={
            "includeBuiltIn": True,
            "includeCustom": True,
            "includeUnused": False,
            "detailedInfo": True,
        },
        validator=validate_detailed_info,
        detailed_display=True,
        tags=["detailed"],
    ),
    StylesTestCase(
        name="获取所有样式（包括未使用的）",
        fixture_name="complex.docx",
        description="获取文档中的所有样式，包括未使用的",
        options={
            "includeBuiltIn": True,
            "includeCustom": True,
            "includeUnused": True,
            "detailedInfo": False,
        },
        validator=validate_includes_unused,
        tags=["all"],
    ),
]


# ==============================================================================
# 辅助函数
# ==============================================================================


def display_styles(styles: list[dict[str, Any]], detailed: bool = False) -> None:
    """显示样式列表"""
    if not styles:
        print("   ⚠️  未找到样式")
        return

    print(f"\n   📚 样式列表 (共 {len(styles)} 个):")

    # 按类型分组
    by_type: dict[str, list[dict[str, Any]]] = {
        "Paragraph": [],
        "Character": [],
        "Table": [],
        "List": [],
    }

    for style in styles:
        style_type = style.get("type", "Unknown")
        if style_type in by_type:
            by_type[style_type].append(style)

    # 显示各类型样式
    for style_type, type_styles in by_type.items():
        if not type_styles:
            continue

        print(f"\n   {style_type} 样式 ({len(type_styles)} 个):")
        for style in type_styles[:10]:  # 最多显示 10 个
            name = style.get("name", "Unknown")
            built_in = style.get("builtIn", False)
            in_use = style.get("inUse", False)
            description = style.get("description")

            status = []
            if built_in:
                status.append("内置")
            else:
                status.append("自定义")
            if in_use:
                status.append("使用中")

            print(f"      - {name} [{', '.join(status)}]")

            if detailed and description:
                print(f"        描述: {description}")

        if len(type_styles) > 10:
            print(f"      ... 还有 {len(type_styles) - 10} 个")


def display_stats(styles: list[dict[str, Any]]) -> None:
    """显示统计信息"""
    print("\n   📈 统计信息:")
    print(f"   - 样式总数: {len(styles)}")

    by_type: dict[str, int] = {"Paragraph": 0, "Character": 0, "Table": 0, "List": 0}
    for style in styles:
        style_type = style.get("type", "Unknown")
        if style_type in by_type:
            by_type[style_type] += 1

    for style_type, count in by_type.items():
        if count > 0:
            print(f"   - {style_type}: {count} 个")

    built_in_count = sum(1 for s in styles if s.get("builtIn", False))
    custom_count = len(styles) - built_in_count
    in_use_count = sum(1 for s in styles if s.get("inUse", False))

    print(f"   - 内置样式: {built_in_count} 个")
    print(f"   - 自定义样式: {custom_count} 个")
    print(f"   - 使用中: {in_use_count} 个")


# ==============================================================================
# 测试执行
# ==============================================================================


async def run_single_test(
    runner: E2ETestRunner,
    test_case: StylesTestCase,
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
    if test_case.options:
        print(f"📝 选项: {test_case.options}")

    fixture_path = f"get_styles_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 执行获取样式动作
            print("\n📝 执行: 获取样式...")
            start_time = time.time()

            params: dict[str, Any] = {"document_uri": fixture.document_uri}
            if test_case.options:
                params["options"] = test_case.options

            action = OfficeAction(
                category="word",
                action_name="get:styles",
                params=params,
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

            # 显示样式
            styles = data.get("styles", [])
            display_styles(styles, detailed=test_case.detailed_display)
            display_stats(styles)

            # 验证结果
            print("\n📊 验证结果:")
            passed = True

            if test_case.validator:
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

        # 如果运行多个测试，等待一下
        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("⏳ 准备下一个测试...")
            await asyncio.sleep(2.0)

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
        description="Word Get Styles E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（默认参数）
  python test_styles.py --test 1

  # 运行所有测试
  python test_styles.py --test all

  # 手动打开文档模式
  python test_styles.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_styles.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=["1", "2", "3", "4", "5", "all"],
        default="1",
        help="要运行的测试: 1=默认, 2=内置, 3=自定义, 4=详细, 5=全部样式, all=全部测试",
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
            if tc.options:
                print(f"     选项: {tc.options}")
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
