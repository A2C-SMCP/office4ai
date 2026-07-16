"""
Basic Export Content E2E Tests (自动化版本)

测试基本的文档导出功能。

测试场景:
1. 纯文本导出 — format=text, 验证 content 非空且包含已知文本
2. HTML 导出 — format=html, 验证 content 包含 HTML 标签
3. Markdown 导出 — format=markdown, 验证 content 非空
4. 空文档导出 — format=text, 验证 success=True, content 为空或极短

运行方式:
    uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test 1
    uv run python manual_tests/word/export_content_e2e/test_basic_export.py --test all
    uv run python manual_tests/word/export_content_e2e/test_basic_export.py --list
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


def validate_text_export(data: dict[str, Any]) -> bool:
    """验证纯文本导出"""
    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False
    if "测试" not in content and "文本" not in content:
        print(f"   ❌ content 不包含预期文本，前 100 字符: {content[:100]}")
        return False
    print(f"   ✅ 纯文本导出成功，content 长度: {len(content)}")
    return True


def validate_html_export(data: dict[str, Any]) -> bool:
    """验证 HTML 导出"""
    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False

    passed = True
    lower = content.lower()

    # 1) 必须包含常见 HTML 结构标签（至少命中一个）
    structural_tags = ["<p", "<div", "<span", "<body", "<html", "<h1", "<h2", "<h3", "<section"]
    found_tags = [t for t in structural_tags if t in lower]
    if not found_tags:
        print(f"   ❌ 未找到任何 HTML 结构标签 ({structural_tags})")
        print(f"      前 300 字符: {content[:300]}")
        passed = False
    else:
        print(f"   ✅ 包含 HTML 结构标签: {found_tags}")

    # 2) 必须有闭合标签（排除自闭合标签的误判）
    import re

    closing_tags = re.findall(r"</\w+>", content)
    if not closing_tags:
        print("   ❌ 未找到任何 HTML 闭合标签 (</...>)")
        passed = False
    else:
        print(f"   ✅ 包含 {len(closing_tags)} 个闭合标签")

    # 3) 导出内容应包含文档中的已知文本
    if "测试" not in content and "文本" not in content:
        print("   ❌ HTML content 不包含预期文本（'测试' 或 '文本'）")
        passed = False
    else:
        print("   ✅ HTML content 包含预期文本")

    print(f"   📏 content 长度: {len(content)}")
    return passed


def validate_markdown_export(data: dict[str, Any]) -> bool:
    """验证 Markdown 导出"""
    content = data.get("content", "")
    if not content:
        print("   ❌ content 为空")
        return False

    passed = True

    # 1) 不应包含 HTML 标签（区分于 HTML 格式）
    import re

    html_tags = re.findall(r"<(?:p|div|span|body|html|table|tr|td)\b", content, re.IGNORECASE)
    if html_tags:
        print(f"   ❌ Markdown 内容不应包含 HTML 标签，但发现: {html_tags[:5]}")
        passed = False
    else:
        print("   ✅ 不包含 HTML 标签（格式正确区分）")

    # 2) 应包含 Markdown 格式特征（至少命中一个）
    md_indicators = {
        "标题 (#)": bool(re.search(r"^#{1,6}\s", content, re.MULTILINE)),
        "段落分隔 (双换行)": "\n\n" in content or "\r\n\r\n" in content,
        "粗体 (**)": "**" in content,
        "列表 (- 或 *)": bool(re.search(r"^[\-\*]\s", content, re.MULTILINE)),
        "有序列表 (1.)": bool(re.search(r"^\d+\.\s", content, re.MULTILINE)),
    }
    found = {k: v for k, v in md_indicators.items() if v}
    if not found:
        print(f"   ⚠️  未检测到常见 Markdown 格式特征: {list(md_indicators.keys())}")
        print(f"      前 300 字符: {content[:300]}")
        # 不强制失败：简单文档可能没有特殊格式
    else:
        print(f"   ✅ 检测到 Markdown 格式特征: {list(found.keys())}")

    # 3) 导出内容应包含文档中的已知文本
    if "测试" not in content and "文本" not in content:
        print("   ❌ Markdown content 不包含预期文本（'测试' 或 '文本'）")
        passed = False
    else:
        print("   ✅ Markdown content 包含预期文本")

    print(f"   📏 content 长度: {len(content)}")
    return passed


def validate_empty_export(data: dict[str, Any]) -> bool:
    """验证空文档导出"""
    content = data.get("content", "")
    # 空文档导出的 content 应该为空或极短（Word 可能有默认段落符）
    if len(content) > 50:
        print(f"   ❌ 空文档 content 过长 ({len(content)} 字符): {content[:100]}")
        return False
    print(f"   ✅ 空文档导出成功，content 长度: {len(content)}")
    return True


# ==============================================================================
# 测试用例
# ==============================================================================

# 每个测试用例对应的导出格式
_EXPORT_FORMATS: list[str] = ["text", "html", "markdown", "text"]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="纯文本导出",
        fixture_name="simple.docx",
        description="以 text 格式导出简单文档，验证 content 非空且包含已知文本",
        validator=validate_text_export,
        tags=["basic"],
    ),
    TestCase(
        name="HTML 导出",
        fixture_name="simple.docx",
        description="以 html 格式导出简单文档，验证 content 包含 HTML 标签",
        validator=validate_html_export,
        tags=["basic"],
    ),
    TestCase(
        name="Markdown 导出",
        fixture_name="simple.docx",
        description="以 markdown 格式导出简单文档，验证 content 非空",
        validator=validate_markdown_export,
        tags=["basic"],
    ),
    TestCase(
        name="空文档导出",
        fixture_name="empty.docx",
        description="以 text 格式导出空文档，验证 success=True 且 content 为空或极短",
        validator=validate_empty_export,
        tags=["edge_case"],
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
    export_format = _EXPORT_FORMATS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            print(f"\n📝 执行: 导出文档内容 (format={export_format})...")
            start_time = time.time()

            action = OfficeAction(
                category="word",
                action_name="export:content",
                params={
                    "document_uri": fixture.document_uri,
                    "format": export_format,
                },
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
                print(f"   content 预览: {content[:100]}{'...' if len(content) > 100 else ''}")

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
        description="Basic Export Content E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=纯文本, 2=HTML, 3=Markdown, 4=空文档, all=全部",
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
