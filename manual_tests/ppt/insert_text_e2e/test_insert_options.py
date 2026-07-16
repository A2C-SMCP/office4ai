"""
PPT Insert Text E2E Tests — Options

测试文本插入的高级选项。

测试场景:
1. 带填充色 — fillColor 选项
2. 连续多个文本框 — 在同一幻灯片连续插入
3. 中文文本 — 中文文本插入验证

运行方式:
    uv run python manual_tests/ppt/insert_text_e2e/test_insert_options.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PPTTestRunner,
    PresentationReader,
    PptTestCase,
    _call_ppt_validator,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_insert_text

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_fill_color(data: dict[str, Any], reader: PresentationReader) -> bool:
    reader.reload()
    if not reader.contains_text("彩色背景"):
        print("   ❌ 未找到 '彩色背景'")
        return False
    print("   ✅ 文档内容验证通过")
    return True


def validate_multiple_insert(data: dict[str, Any], reader: PresentationReader) -> bool:
    reader.reload()
    if not reader.contains_text("第二个文本框"):
        print("   ❌ 未找到 '第二个文本框'")
        return False
    print("   ✅ 文档内容验证通过: 找到连续插入的文本")
    return True


def validate_chinese_text(data: dict[str, Any], reader: PresentationReader) -> bool:
    reader.reload()
    if not reader.contains_text("中文测试文本"):
        print("   ❌ 未找到 '中文测试文本'")
        return False
    print("   ✅ 文档内容验证通过: 中文文本完整")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="带填充色",
        fixture_name="empty.pptx",
        description="插入带黄色背景的文本框",
        validator=validate_fill_color,
        tags=["advanced"],
    ),
    PptTestCase(
        name="连续多个文本框",
        fixture_name="empty.pptx",
        description="在同一幻灯片连续插入两个文本框",
        validator=validate_multiple_insert,
        tags=["advanced"],
    ),
    PptTestCase(
        name="中文文本插入",
        fixture_name="empty.pptx",
        description="插入较长的中文文本",
        validator=validate_chinese_text,
        tags=["basic"],
    ),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()

            if test_number == 1:
                # 带填充色
                success, data, error = await ppt_insert_text(
                    workspace,
                    fixture.document_uri,
                    "彩色背景",
                    options={"fillColor": "#FFFF00", "left": 100, "top": 100},
                )
            elif test_number == 2:
                # 连续多个
                success1, _, error1 = await ppt_insert_text(
                    workspace, fixture.document_uri, "第一个文本框", options={"left": 50, "top": 50}
                )
                if not success1:
                    print(f"❌ 第一次插入失败: {error1}")
                    return False
                success, data, error = await ppt_insert_text(
                    workspace, fixture.document_uri, "第二个文本框", options={"left": 50, "top": 200}
                )
            else:
                # 中文文本
                success, data, error = await ppt_insert_text(
                    workspace,
                    fixture.document_uri,
                    "中文测试文本：这是一段用于验证 PPT 中文文本插入的长文本。确保中文字符不会乱码。",
                )

            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 插入失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}

            print("\n📊 验证结果:")
            passed = True
            if test_case.validator:
                await asyncio.sleep(0.5)
                reader = PresentationReader(fixture.working_path)
                if not _call_ppt_validator(test_case.validator, data, reader):
                    passed = False

            print("\n" + "=" * 70)
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
            print("=" * 70)
            return passed

    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        import traceback

        traceback.print_exc()
        return False


async def run_tests(test_indices: list[int], auto_open: bool = True, cleanup_on_success: bool = True) -> bool:
    ensure_ppt_fixtures(FIXTURES_DIR)
    runner = PPTTestRunner(fixtures_dir=FIXTURES_DIR.parent, auto_open=auto_open, cleanup_on_success=cleanup_on_success)
    results: list[bool] = []
    for idx in test_indices:
        if idx < 1 or idx > len(TEST_CASES):
            continue
        if len(test_indices) > 1 and results:
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续...")
        result = await run_single_test(runner, TEST_CASES[idx - 1], idx)
        results.append(result)
    if len(results) > 1:
        print(f"\n📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    return all(results)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="PPT Insert Text Options E2E Tests")
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--always-cleanup", action="store_true")
    parser.add_argument("--list", action="store_true")
    args = parser.parse_args()

    if args.list:
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name} — {tc.description}")
        return

    test_indices = list(range(1, len(TEST_CASES) + 1)) if args.test == "all" else [int(args.test)]
    try:
        success = asyncio.run(
            run_tests(test_indices, auto_open=not args.no_auto_open, cleanup_on_success=not args.always_cleanup)
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
