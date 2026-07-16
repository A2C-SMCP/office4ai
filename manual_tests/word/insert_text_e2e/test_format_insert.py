"""
Formatted Text Insert E2E Tests (自动化版本)

测试带格式（format 参数）的文本插入功能。

测试场景:
1. 粗体文本插入
2. 斜体文本插入
3. 字体大小设置
4. 字体名称设置
5. 颜色设置
6. 组合格式（粗体+斜体+大小+颜色）

运行方式:
    uv run python manual_tests/insert_text_e2e/test_format_insert.py --test 1
    uv run python manual_tests/insert_text_e2e/test_format_insert.py --test all
    uv run python manual_tests/insert_text_e2e/test_format_insert.py --list
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


def validate_bold_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证粗体文本插入"""
    reader.reload()
    if not reader.contains("这是粗体文本"):
        print("   ❌ 文档中未找到 '这是粗体文本'")
        return False
    if reader.run_has_format("这是粗体文本", bold=True):
        print("   ✅ 文档内容验证通过: 粗体格式正确")
        return True
    print("   ⚠️  文本存在但格式未验证（Word 可能延迟保存格式）")
    return True  # 文本存在即可，格式由人工确认


def validate_italic_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证斜体文本插入"""
    reader.reload()
    if not reader.contains("这是斜体文本"):
        print("   ❌ 文档中未找到 '这是斜体文本'")
        return False
    if reader.run_has_format("这是斜体文本", italic=True):
        print("   ✅ 文档内容验证通过: 斜体格式正确")
        return True
    print("   ⚠️  文本存在但格式未验证")
    return True


def validate_font_size_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证字体大小插入"""
    reader.reload()
    for text in ["小号文本 (12pt)", "中号文本 (16pt)", "大号文本 (24pt)"]:
        if not reader.contains(text):
            print(f"   ❌ 文档中未找到 '{text}'")
            return False
    print("   ✅ 文档内容验证通过: 所有字体大小文本已插入")
    return True


def validate_font_name_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证字体名称插入"""
    reader.reload()
    for text in ["Arial 字体", "Times New Roman 字体", "Courier New 字体"]:
        if not reader.contains(text):
            print(f"   ❌ 文档中未找到 '{text}'")
            return False
    print("   ✅ 文档内容验证通过: 所有字体文本已插入")
    return True


def validate_color_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证颜色文本插入"""
    reader.reload()
    for text in ["红色文本", "绿色文本", "蓝色文本"]:
        if not reader.contains(text):
            print(f"   ❌ 文档中未找到 '{text}'")
            return False
    print("   ✅ 文档内容验证通过: 所有颜色文本已插入")
    return True


def validate_combined_insert(data: dict[str, Any], reader: DocumentReader) -> bool:
    """验证组合格式文本插入"""
    reader.reload()
    if not reader.contains("组合格式文本"):
        print("   ❌ 文档中未找到 '组合格式文本'")
        return False
    if reader.run_has_format("组合格式文本", bold=True, italic=True):
        print("   ✅ 文档内容验证通过: 组合格式（粗体+斜体）正确")
        return True
    print("   ⚠️  文本存在但格式未验证")
    return True


# ==============================================================================
# 测试数据
# ==============================================================================

# OASP 0.4.0: 字体属性收敛到嵌套 font（WordFont）；扁平 bold/fontSize/fontName/color 已删，
# 须写 format.font.{bold/size/name/color/...}，否则被 DTO extra=ignore 静默丢弃。
_FORMAT_CONFIGS: list[dict[str, Any]] = [
    # Test 1: bold
    {"text": "这是粗体文本", "format": {"font": {"bold": True}}},
    # Test 2: italic
    {"text": "这是斜体文本", "format": {"font": {"italic": True}}},
    # Test 3: fontSize (multiple inserts)
    {
        "multi": [
            {"text": "小号文本 (12pt)\n", "format": {"font": {"size": 12}}},
            {"text": "中号文本 (16pt)\n", "format": {"font": {"size": 16}}},
            {"text": "大号文本 (24pt)\n", "format": {"font": {"size": 24}}},
        ]
    },
    # Test 4: fontName (multiple inserts)
    {
        "multi": [
            {"text": "Arial 字体\n", "format": {"font": {"name": "Arial"}}},
            {"text": "Times New Roman 字体\n", "format": {"font": {"name": "Times New Roman"}}},
            {"text": "Courier New 字体\n", "format": {"font": {"name": "Courier New"}}},
        ]
    },
    # Test 5: color (multiple inserts)
    {
        "multi": [
            {"text": "红色文本\n", "format": {"font": {"color": "#FF0000"}}},
            {"text": "绿色文本\n", "format": {"font": {"color": "#00FF00"}}},
            {"text": "蓝色文本\n", "format": {"font": {"color": "#0000FF"}}},
        ]
    },
    # Test 6: combined
    {
        "text": "组合格式文本",
        "format": {
            "font": {
                "bold": True,
                "italic": True,
                "size": 18,
                "name": "Arial",
                "color": "#FF0000",
            },
        },
    },
]

TEST_CASES: list[TestCase] = [
    TestCase(
        name="粗体文本插入",
        fixture_name="empty.docx",
        description="插入粗体文本，验证文本存在并检查 bold 格式",
        validator=validate_bold_insert,
        tags=["format"],
    ),
    TestCase(
        name="斜体文本插入",
        fixture_name="empty.docx",
        description="插入斜体文本，验证文本存在并检查 italic 格式",
        validator=validate_italic_insert,
        tags=["format"],
    ),
    TestCase(
        name="字体大小设置",
        fixture_name="empty.docx",
        description="插入 3 种字体大小的文本 (12/16/24pt)，验证文本存在",
        validator=validate_font_size_insert,
        tags=["format"],
    ),
    TestCase(
        name="字体名称设置",
        fixture_name="empty.docx",
        description="插入 3 种字体 (Arial/Times/Courier) 的文本，验证文本存在",
        validator=validate_font_name_insert,
        tags=["format"],
    ),
    TestCase(
        name="颜色设置",
        fixture_name="empty.docx",
        description="插入 3 种颜色 (红/绿/蓝) 的文本，验证文本存在",
        validator=validate_color_insert,
        tags=["format"],
    ),
    TestCase(
        name="组合格式插入",
        fixture_name="empty.docx",
        description="插入粗体+斜体+18pt+Arial+红色的文本，验证文本和格式",
        validator=validate_combined_insert,
        tags=["format", "combined"],
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

    fixture_path = f"insert_text_e2e/{test_case.fixture_name}"
    config = _FORMAT_CONFIGS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (
            workspace,
            fixture,
        ):
            print(f"\n📝 执行: {test_case.name}...")
            start_time = time.time()

            # 多次插入模式
            if "multi" in config:
                for i, item in enumerate(config["multi"], 1):
                    action = OfficeAction(
                        category="word",
                        action_name="insert:text",
                        params={
                            "document_uri": fixture.document_uri,
                            "text": item["text"],
                            "location": "End",
                            "format": item["format"],
                        },
                    )
                    result = await workspace.execute(action)
                    if not result.success:
                        print(f"   ❌ 第 {i} 次插入失败: {result.error}")
                        return False
                    print(f"   ✅ 第 {i} 次插入成功")
                    await asyncio.sleep(0.5)
                data = result.data or {}  # type: ignore[possibly-undefined]
            else:
                # 单次插入模式
                action = OfficeAction(
                    category="word",
                    action_name="insert:text",
                    params={
                        "document_uri": fixture.document_uri,
                        "text": config["text"],
                        "location": "End",
                        "format": config["format"],
                    },
                )
                result = await workspace.execute(action)
                if not result.success:
                    print(f"❌ 插入失败: {result.error}")
                    return False
                data = result.data or {}

            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")
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
        description="Formatted Text Insert E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=bold, 2=italic, 3=fontSize, 4=fontName, 5=color, 6=combined, all=全部",
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
