"""
Edge Cases Get Visible Content E2E Tests (E2ETestRunner 版本)

测试 word:get:visibleContent 的边界情况和异常场景。

特性：
- 自动复制测试文档
- 自动打开 Word
- 自动验证结果
- 成功后自动清理，失败则保留供调试

测试场景:
1. 超长文档 — 验证大文档可见内容获取
2. 特殊字符内容 — 验证文本非空
3. 嵌入对象 — 验证含图片/表格的文档有元素返回
4. 连续请求一致性 — 连续 3 次获取，验证文本一致

运行方式:
    # 运行单个测试
    uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 1

    # 运行所有测试
    uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test all

    # 不自动打开文档（手动打开）
    uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 1 --no-auto-open

    # 失败时也清理文件
    uv run python manual_tests/get_visible_content_e2e/test_edge_cases.py --test 1 --always-cleanup
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


def validate_large_document(data: dict[str, Any]) -> bool:
    """
    验证超长文档

    检查:
    - text 长度 > 100 字符
    """
    text = data.get("text", "")
    if len(text) <= 100:
        print(f"      ⚠️  大文档文本过短: {len(text)} 字符 (预期 > 100)")
        return False

    print(f"      ✅ 大文档文本长度: {len(text)} 字符 (> 100)")
    return True


def validate_special_characters(data: dict[str, Any]) -> bool:
    """
    验证特殊字符内容

    检查:
    - text 字段非空
    """
    text = data.get("text")
    if not text:
        print("      ⚠️  text 字段为空或不存在")
        return False

    print(f"      ✅ text 非空，长度: {len(text)}")
    return True


def validate_embedded_objects(data: dict[str, Any]) -> bool:
    """
    验证嵌入对象

    检查:
    - elements 列表非空
    """
    elements = data.get("elements", [])
    if len(elements) <= 0:
        print("      ⚠️  elements 列表为空")
        return False

    # 统计元素类型
    elem_types: dict[str, int] = {}
    for elem in elements:
        elem_type = elem.get("type", "unknown")
        elem_types[elem_type] = elem_types.get(elem_type, 0) + 1

    type_summary = ", ".join(f"{t}: {c}" for t, c in elem_types.items())
    print(f"      ✅ elements 数量: {len(elements)} ({type_summary})")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[VisibleContentTestCase] = [
    VisibleContentTestCase(
        name="超长文档",
        fixture_name="large.docx",
        description="获取大文档的可见内容（maxTextLength=10000），验证文本长度 > 100",
        options={"maxTextLength": 10000},
        validator=validate_large_document,
        tags=["edge"],
    ),
    VisibleContentTestCase(
        name="特殊字符内容",
        fixture_name="simple.docx",
        description="获取包含特殊字符的文档内容，验证 text 非空",
        options=None,
        validator=validate_special_characters,
        tags=["edge"],
    ),
    VisibleContentTestCase(
        name="嵌入对象",
        fixture_name="complex.docx",
        description="获取含嵌入对象（图片、表格）的文档，验证 elements 非空",
        options={"includeText": True, "includeImages": True, "includeTables": True},
        validator=validate_embedded_objects,
        tags=["edge"],
    ),
    VisibleContentTestCase(
        name="连续请求一致性",
        fixture_name="simple.docx",
        description="连续 3 次获取可见内容，验证每次返回的 text 一致",
        options=None,
        validator=None,  # 使用自定义执行逻辑
        tags=["edge", "consistency"],
    ),
]


# ==============================================================================
# 测试执行
# ==============================================================================


async def _execute_consecutive_requests(
    runner: E2ETestRunner,
    test_case: VisibleContentTestCase,
    test_number: int,
) -> bool:
    """
    测试 4 专用执行器：连续请求一致性

    在同一个 workspace 连接中执行 3 次 get:visibleContent，
    比较每次返回的 text 是否一致。

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

    fixture_path = f"get_visible_content_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            texts: list[str] = []
            request_count = 3

            for i in range(1, request_count + 1):
                print(f"\n--- 第 {i}/{request_count} 次请求 ---")
                start_time = time.time()

                action = OfficeAction(
                    category="word",
                    action_name="get:visibleContent",
                    params={"document_uri": fixture.document_uri},
                )

                result = await workspace.execute(action)
                elapsed_ms = (time.time() - start_time) * 1000

                print(f"   ⏱️  执行时间: {elapsed_ms:.1f}ms")

                if not result.success:
                    print(f"   ❌ 第 {i} 次获取失败: {result.error}")
                    return False

                data = result.data or {}
                text = data.get("text", "")
                texts.append(text)
                print(f"   ✅ 获取成功，文本长度: {len(text)}")

                # 请求间隔
                if i < request_count:
                    await asyncio.sleep(1.0)

            # 验证一致性
            print("\n📊 验证一致性:")
            passed = True

            if len(set(texts)) == 1:
                print(f"   ✅ {request_count} 次请求返回的文本完全一致 (长度: {len(texts[0])})")
            else:
                print(f"   ❌ {request_count} 次请求返回的文本不一致!")
                for i, text in enumerate(texts, 1):
                    print(f"      第 {i} 次: 长度={len(text)}, 前50字符={repr(text[:50])}")
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


async def run_single_test(
    runner: E2ETestRunner,
    test_case: VisibleContentTestCase,
    test_number: int,
) -> bool:
    """
    执行单个测试用例

    测试 4（连续请求一致性）使用自定义执行逻辑，
    其他测试使用标准流程。

    Args:
        runner: E2E 测试运行器
        test_case: 测试用例
        test_number: 测试编号

    Returns:
        是否通过
    """
    # 测试 4: 连续请求一致性 — 使用专用执行器
    if test_number == 4:
        return await _execute_consecutive_requests(runner, test_case, test_number)

    # 标准执行流程
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
        description="Get Visible Content E2E Tests - Edge Cases (E2ETestRunner 版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 运行测试 1（超长文档）
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
        help="要运行的测试: 1=超长文档, 2=特殊字符, 3=嵌入对象, 4=连续请求, all=全部",
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
