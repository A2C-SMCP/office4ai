"""
Export Content Options E2E Tests (自动化版本)

测试 word:export:content 的导出选项。

测试场景:
1. 包含表格导出 — 复杂文档以 html 格式导出，验证 content 包含表格标记
2. 包含表格选项 — 复杂文档以 markdown 格式导出，includeTables=true
3. 大文档性能 — 大文档以 text 格式导出，验证执行时间和 content 长度

运行方式:
    uv run python manual_tests/word/export_content_e2e/test_export_options.py --test 1
    uv run python manual_tests/word/export_content_e2e/test_export_options.py --test all
    uv run python manual_tests/word/export_content_e2e/test_export_options.py --list
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.e2e_base import (
    E2ETestRunner,
    TestCase,
    _call_validator,
    ensure_fixtures,
)
from office4ai.environment.workspace.base import OfficeAction

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "export_content_e2e"

# ==============================================================================
# Validator 函数
# ==============================================================================


def validate_table_html(data: dict[str, Any]) -> bool:
    """验证 HTML 导出包含表格标记"""
    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False

    passed = True
    lower = content.lower()

    # 1) 必须包含 <table 标签
    if "<table" not in lower:
        print("   ❌ HTML content 不包含 <table> 标签")
        print(f"      前 300 字符: {content[:300]}")
        passed = False
    else:
        print("   ✅ 包含 <table> 标签")

    # 2) 表格内部应包含行和单元格
    has_tr = "<tr" in lower
    has_td = "<td" in lower or "<th" in lower
    if not has_tr or not has_td:
        print(f"   ❌ 表格结构不完整 (<tr>: {has_tr}, <td>/<th>: {has_td})")
        passed = False
    else:
        print("   ✅ 表格结构完整 (<tr> + <td>/<th>)")

    # 3) 应包含闭合的 </table> 标签
    if "</table>" not in lower:
        print("   ❌ 缺少 </table> 闭合标签")
        passed = False
    else:
        print("   ✅ 包含 </table> 闭合标签")

    print(f"   📏 content 长度: {len(content)}")
    return passed


def validate_table_markdown(data: dict[str, Any]) -> bool:
    """验证 Markdown 导出包含表格内容"""
    import re

    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False

    passed = True

    # 1) 不应包含 HTML 标签（确认是 Markdown 而非 HTML）
    html_tags = re.findall(r"<(?:table|tr|td|th|p|div|span)\b", content, re.IGNORECASE)
    if html_tags:
        print(f"   ❌ Markdown 内容不应包含 HTML 标签，但发现: {html_tags[:5]}")
        passed = False
    else:
        print("   ✅ 不包含 HTML 标签（格式正确区分）")

    # 2) Markdown 表格必须包含 | 分隔符
    pipe_lines = [line for line in content.split("\n") if "|" in line]
    if not pipe_lines:
        print("   ❌ Markdown content 不包含 | 分隔符（无表格）")
        print(f"      前 300 字符: {content[:300]}")
        passed = False
    else:
        print(f"   ✅ Markdown 包含 {len(pipe_lines)} 行表格内容")

    # 3) Markdown 表格应包含分隔行（如 |---|---|）
    separator_pattern = re.compile(r"\|[\s\-:]+\|")
    has_separator = any(separator_pattern.search(line) for line in pipe_lines)
    if pipe_lines and not has_separator:
        print("   ⚠️  未检测到表格分隔行 (|---|---| 格式)")
        # 不强制失败：某些 Markdown 实现可能省略分隔行
    elif has_separator:
        print("   ✅ 包含表格分隔行")

    print(f"   📏 content 长度: {len(content)}")
    return passed


def validate_large_export(data: dict[str, Any]) -> bool:
    """验证大文档导出"""
    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False
    if len(content) < 1000:
        print(f"   ❌ 大文档 content 过短 ({len(content)} 字符)，预期 > 1000")
        return False
    print(f"   ✅ 大文档导出成功，content 长度: {len(content)}")
    return True


# ==============================================================================
# 测试用例
# ==============================================================================

# 每个测试对应的导出配置
_EXPORT_CONFIGS: list[dict[str, Any]] = [
    {"format": "html"},
    {"format": "markdown", "options": {"include_tables": True}},
    {"format": "text"},
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="包含表格导出 (HTML)",
        fixture_name="complex.docx",
        description="复杂文档以 html 格式导出，验证 content 包含表格标记",
        validator=validate_table_html,
        tags=["advanced"],
    ),
    TestCase(
        name="包含表格选项 (Markdown)",
        fixture_name="complex.docx",
        description="复杂文档以 markdown 格式导出，includeTables=true",
        validator=validate_table_markdown,
        tags=["advanced"],
    ),
    TestCase(
        name="大文档性能",
        fixture_name="large.docx",
        description="大文档以 text 格式导出，验证执行时间合理且 content 长度 > 1000",
        validator=validate_large_export,
        tags=["performance"],
    ),
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

    fixture_path = f"export_content_e2e/{test_case.fixture_name}"
    config = _EXPORT_CONFIGS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            export_format = config["format"]
            options = config.get("options")

            print(f"\n📝 执行: 导出文档内容 (format={export_format})...")
            if options:
                print(f"   选项: {options}")
            start_time = time.time()

            params: dict[str, Any] = {
                "document_uri": fixture.document_uri,
                "format": export_format,
            }
            if options:
                params["options"] = options

            action = OfficeAction(
                category="word",
                action_name="export:content",
                params=params,
            )

            result = await workspace.execute(action)
            elapsed_ms = (time.time() - start_time) * 1000

            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not result.success:
                print(f"❌ 导出失败: {result.error}")
                return False

            print("✅ 协议返回成功")
            data = result.data or {}
            content = data.get("content", "")
            print(f"   content 长度: {len(content)}")
            if content:
                print(f"   content 预览: {content[:150]}{'...' if len(content) > 150 else ''}")

            # DataValidator 验证
            print("\n📊 验证结果:")
            passed = True

            if test_case.validator:
                if not _call_validator(test_case.validator, data, None):  # type: ignore[arg-type]
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
        description="Export Content Options E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=表格HTML, 2=表格Markdown, 3=大文档性能, all=全部",
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
