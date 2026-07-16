"""
Options Get Visible Content E2E Tests (E2ETestRunner 版本)

测试 word:get:visibleContent 的各种选项参数组合。

特性：
- 自动复制测试文档
- 自动打开 Word
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. includeText + includeImages — 排除表格元素
2. includeText + includeTables — 排除图片元素
3. maxTextLength — 限制单个文本元素长度
4. text only — 排除图片和表格元素
5. detailedMetadata — 获取详细元数据

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_visible_content_e2e/test_options_get.py --test 1 --always-cleanup
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
    _call_validator,
    ensure_fixtures,
)
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

# 夹具目录
FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "get_visible_content_e2e"


@dataclass
class VisibleContentTestCase:
    """
    可见内容测试用例定义

    Attributes:
        name: 测试名称
        fixture_name: 夹具文件名（相对于 fixtures 目录）
        description: 测试描述
        options: 获取可见内容的选项
        validator: 自定义验证函数 (DataValidator: data -> bool)
        tags: 标签列表
    """

    name: str
    fixture_name: str
    description: str
    options: dict[str, Any] | None = None
    validator: Any | None = None
    tags: list[str] = field(default_factory=list)


# ==============================================================================
# 验证函数 (DataValidator: data -> bool)
# ==============================================================================


def validate_no_table_elements(data: dict[str, Any]) -> bool:
    """
    验证不包含 table 元素

    检查:
    - elements 中没有 type=table 的元素
    """
    elements = data.get("elements", [])
    table_elements = [e for e in elements if e.get("type") == "table"]
    if len(table_elements) > 0:
        print(f"      ⚠️  includeTables=False 但返回了 {len(table_elements)} 个 table 元素")
        return False

    # 统计实际元素
    elem_types: dict[str, int] = {}
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_types[elem_type] = elem_types.get(elem_type, 0) + 1

    type_summary = ", ".join(f"{t}: {c}" for t, c in elem_types.items())
    print(f"      ✅ 无 table 元素，实际元素: {type_summary}")
    return True


def validate_no_image_elements(data: dict[str, Any]) -> bool:
    """
    验证不包含 image 元素

    检查:
    - elements 中没有 type=image 的元素
    """
    elements = data.get("elements", [])
    image_elements = [e for e in elements if e.get("type") == "image"]
    if len(image_elements) > 0:
        print(f"      ⚠️  includeImages=False 但返回了 {len(image_elements)} 个 image 元素")
        return False

    # 统计实际元素
    elem_types: dict[str, int] = {}
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_types[elem_type] = elem_types.get(elem_type, 0) + 1

    type_summary = ", ".join(f"{t}: {c}" for t, c in elem_types.items())
    print(f"      ✅ 无 image 元素，实际元素: {type_summary}")
    return True


def validate_max_text_length(data: dict[str, Any]) -> bool:
    """
    验证 maxTextLength 限制

    检查:
    - 所有 text 类型元素的 content.text 长度 <= 100
    """
    max_length = 100
    elements = data.get("elements", [])
    text_elements = [e for e in elements if e.get("type") == "text"]

    exceeded: list[str] = []
    for i, elem in enumerate(text_elements):
        content = elem.get("content", {})
        elem_text = content.get("text", "")
        if len(elem_text) > max_length:
            exceeded.append(f"元素#{i + 1}: {len(elem_text)} 字符")

    if exceeded:
        print(f"      ⚠️  以下元素超过 maxTextLength={max_length}:")
        for item in exceeded:
            print(f"         - {item}")
        return False

    print(f"      ✅ 所有 {len(text_elements)} 个 text 元素长度 <= {max_length}")
    return True


def validate_text_only(data: dict[str, Any]) -> bool:
    """
    验证仅返回 text 元素（无 image/table）

    检查:
    - elements 中没有 type=image 或 type=table 的元素
    """
    elements = data.get("elements", [])
    image_elements = [e for e in elements if e.get("type") == "image"]
    table_elements = [e for e in elements if e.get("type") == "table"]

    if len(image_elements) > 0:
        print(f"      ⚠️  includeImages=False 但返回了 {len(image_elements)} 个 image 元素")
        return False

    if len(table_elements) > 0:
        print(f"      ⚠️  includeTables=False 但返回了 {len(table_elements)} 个 table 元素")
        return False

    text_count = sum(1 for e in elements if e.get("type") == "text")
    print(f"      ✅ 仅 text 元素: {text_count} 个 (无 image/table)")
    return True


def validate_detailed_metadata(data: dict[str, Any]) -> bool:
    """
    验证 detailedMetadata 模式

    检查:
    - elements 列表非空（有元素返回即可）
    """
    elements = data.get("elements", [])
    if len(elements) == 0:
        print("      ⚠️  elements 列表为空")
        return False

    # 统计元素类型
    elem_types: dict[str, int] = {}
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_types[elem_type] = elem_types.get(elem_type, 0) + 1

    type_summary = ", ".join(f"{t}: {c}" for t, c in elem_types.items())
    print(f"      ✅ detailedMetadata 返回 {len(elements)} 个元素 ({type_summary})")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[VisibleContentTestCase] = [
    VisibleContentTestCase(
        name="includeText + includeImages",
        fixture_name="complex.docx",
        description="包含文本和图片，排除表格（includeTables=False），验证无 table 元素",
        options={"includeText": True, "includeImages": True, "includeTables": False},
        validator=validate_no_table_elements,
        tags=["options"],
    ),
    VisibleContentTestCase(
        name="includeText + includeTables",
        fixture_name="complex.docx",
        description="包含文本和表格，排除图片（includeImages=False），验证无 image 元素",
        options={"includeText": True, "includeImages": False, "includeTables": True},
        validator=validate_no_image_elements,
        tags=["options"],
    ),
    VisibleContentTestCase(
        name="maxTextLength",
        fixture_name="complex.docx",
        description="限制文本元素最大长度为 100 字符，验证所有 text 元素不超限",
        options={"includeText": True, "maxTextLength": 100},
        validator=validate_max_text_length,
        tags=["options"],
    ),
    VisibleContentTestCase(
        name="text only",
        fixture_name="complex.docx",
        description="仅获取文本，排除图片和表格，验证无 image/table 元素",
        options={"includeText": True, "includeImages": False, "includeTables": False},
        validator=validate_text_only,
        tags=["options"],
    ),
    VisibleContentTestCase(
        name="detailedMetadata",
        fixture_name="complex.docx",
        description="使用 detailedMetadata=True 获取详细元数据，验证 elements 非空",
        options={"includeText": True, "detailedMetadata": True},
        validator=validate_detailed_metadata,
        tags=["options"],
    ),
]


# ==============================================================================
# 测试执行
# ==============================================================================


async def run_single_test(
    runner: E2ETestRunner,
    test_case: VisibleContentTestCase,
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

    fixture_path = f"get_visible_content_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            # 构造 action
            print("\n📝 执行: 获取可见内容...")
            start_time = time.time()

            options = test_case.options
            action = OfficeAction(
                category="word",
                action_name="get:visibleContent",
                params={
                    "document_uri": fixture.document_uri,
                    **({"options": options} if options else {}),
                },
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

            # 显示摘要
            text = data.get("text", "")
            elements = data.get("elements", [])
            metadata = data.get("metadata", {})
            print(f"   文本长度: {len(text)} 字符")
            print(f"   元素数量: {len(elements)}")
            print(f"   字符数(metadata): {metadata.get('characterCount', 'N/A')}")

            # 元素类型统计
            elem_types: dict[str, int] = {}
            for elem in elements:
                elem_type = elem.get("type", "unknown")
                elem_types[elem_type] = elem_types.get(elem_type, 0) + 1
            if elem_types:
                print(f"   元素类型: {elem_types}")

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
        description="Get Visible Content E2E Tests - Options (E2ETestRunner 版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（includeText + includeImages）
  python test_options_get.py --test 1

  # 运行所有测试
  python test_options_get.py --test all

  # 手动打开文档模式
  python test_options_get.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_options_get.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=text+images, 2=text+tables, 3=maxTextLength, 4=text only, 5=detailedMetadata, all=全部",
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
