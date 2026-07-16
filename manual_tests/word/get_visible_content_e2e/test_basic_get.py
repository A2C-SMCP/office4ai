"""
Basic Get Visible Content E2E Tests (E2ETestRunner 版本)

测试 word:get:visibleContent 的基础功能。

特性：
- 自动复制测试文档
- 自动打开 Word
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. 获取文本内容 — 简单文档，默认参数
2. 获取空文档 — 空文档应返回极少字符
3. 获取格式化文本 — 带 detailedMetadata 选项
4. 获取多种元素 — 文本 + 图片 + 表格

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_visible_content_e2e/test_basic_get.py --test 1 --always-cleanup
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


def validate_text_content(data: dict[str, Any]) -> bool:
    """
    验证基本文本内容

    检查:
    - text 字段存在且非空
    - elements 列表存在且至少有一个元素
    """
    text = data.get("text")
    if not text:
        print("      ⚠️  text 字段为空或不存在")
        return False

    elements = data.get("elements", [])
    if len(elements) <= 0:
        print("      ⚠️  elements 列表为空")
        return False

    print(f"      ✅ text 长度: {len(text)}, elements 数量: {len(elements)}")
    return True


def validate_empty_document(data: dict[str, Any]) -> bool:
    """
    验证空文档

    检查:
    - 字符数 <= 10（Word 空文档通常有少量隐藏字符）
    """
    char_count = data.get("metadata", {}).get("characterCount", 0)
    if char_count > 10:
        print(f"      ⚠️  空文档字符数过多: {char_count} (阈值: 10)")
        return False

    print(f"      ✅ 空文档字符数: {char_count} <= 10")
    return True


def validate_formatted_text(data: dict[str, Any]) -> bool:
    """
    验证格式化文本

    检查:
    - elements 中至少有一个 text 类型的元素
    """
    elements = data.get("elements", [])
    text_elements = [e for e in elements if e.get("type") == "text"]
    if len(text_elements) == 0:
        print("      ⚠️  未找到 text 类型的元素")
        return False

    print(f"      ✅ 找到 {len(text_elements)} 个 text 类型元素")
    return True


def validate_mixed_elements(data: dict[str, Any]) -> bool:
    """
    验证多种元素

    检查:
    - 同时包含 text 和 table 类型的元素
    """
    elements = data.get("elements", [])
    text_elements = [e for e in elements if e.get("type") == "text"]
    table_elements = [e for e in elements if e.get("type") == "table"]

    if len(text_elements) == 0:
        print("      ⚠️  未找到 text 类型的元素")
        return False

    if len(table_elements) == 0:
        print("      ⚠️  未找到 table 类型的元素")
        return False

    print(f"      ✅ text: {len(text_elements)} 个, table: {len(table_elements)} 个")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[VisibleContentTestCase] = [
    VisibleContentTestCase(
        name="获取文本内容",
        fixture_name="simple.docx",
        description="获取简单文档的可见内容，验证 text 非空且 elements 列表非空",
        options=None,
        validator=validate_text_content,
        tags=["basic"],
    ),
    VisibleContentTestCase(
        name="获取空文档",
        fixture_name="empty.docx",
        description="获取空文档的可见内容，字符数应 <= 10（Word 空文档阈值）",
        options=None,
        validator=validate_empty_document,
        tags=["basic"],
    ),
    VisibleContentTestCase(
        name="获取格式化文本",
        fixture_name="complex.docx",
        description="使用 detailedMetadata + includeText 获取格式化文本，验证有 text 类型元素",
        options={"detailedMetadata": True, "includeText": True},
        validator=validate_formatted_text,
        tags=["basic"],
    ),
    VisibleContentTestCase(
        name="获取多种元素",
        fixture_name="complex.docx",
        description="同时获取文本、图片、表格，验证包含 text 和 table 类型元素",
        options={"includeText": True, "includeImages": True, "includeTables": True},
        validator=validate_mixed_elements,
        tags=["basic"],
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
        description="Get Visible Content E2E Tests - Basic (E2ETestRunner 版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（获取文本内容）
  python test_basic_get.py --test 1

  # 运行所有测试
  python test_basic_get.py --test all

  # 手动打开文档模式
  python test_basic_get.py --test 1 --no-auto-open

  # 失败时也清理文件
  python test_basic_get.py --test 1 --always-cleanup
        """,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=文本内容, 2=空文档, 3=格式化文本, 4=多种元素, all=全部",
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
