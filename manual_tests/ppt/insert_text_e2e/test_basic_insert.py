"""
PPT Insert Text E2E Tests — Basic

测试基本的 PPT 文本插入功能。

测试场景:
1. 简单文本插入 — 在空白幻灯片插入简单文本
2. 带位置插入 — 指定 left/top/width/height
3. 带字体插入 — 指定 fontSize/fontName
4. 指定幻灯片插入 — 在 slideIndex=1 上插入

运行方式:
    uv run python manual_tests/ppt/insert_text_e2e/test_basic_insert.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PptTestCase,
    PPTTestRunner,
    PresentationReader,
    _call_ppt_validator,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_insert_text

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_simple_insert(data: dict[str, Any], reader: PresentationReader) -> bool:
    """验证简单文本插入"""
    reader.reload()
    if not reader.contains_text("Hello PPT"):
        print("   ❌ 演示文稿中未找到 'Hello PPT'")
        return False
    print("   ✅ 文档内容验证通过: 包含 'Hello PPT'")
    return True


def validate_positioned_insert(data: dict[str, Any], reader: PresentationReader) -> bool:
    """验证带位置的文本插入"""
    reader.reload()
    if not reader.contains_text("定位文本"):
        print("   ❌ 演示文稿中未找到 '定位文本'")
        return False
    print("   ✅ 文档内容验证通过: 包含 '定位文本'")
    return True


def validate_font_insert(data: dict[str, Any], reader: PresentationReader) -> bool:
    """验证带字体的文本插入"""
    reader.reload()
    if not reader.contains_text("字体测试"):
        print("   ❌ 演示文稿中未找到 '字体测试'")
        return False
    print("   ✅ 文档内容验证通过: 包含 '字体测试'")
    return True


def validate_slide_index_insert(data: dict[str, Any], reader: PresentationReader) -> bool:
    """验证指定幻灯片的文本插入"""
    reader.reload()
    if not reader.slide_has_text(1, "第二页文本"):
        print("   ❌ 第二张幻灯片中未找到 '第二页文本'")
        return False
    print("   ✅ 文档内容验证通过: 第二张幻灯片包含 '第二页文本'")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="简单文本插入",
        fixture_name="empty.pptx",
        description="在空白幻灯片插入 'Hello PPT'",
        validator=validate_simple_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带位置插入",
        fixture_name="empty.pptx",
        description="在指定位置 (100, 100) 插入文本",
        validator=validate_positioned_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带字体插入",
        fixture_name="empty.pptx",
        description="插入带 fontSize=24, fontName=Arial 的文本",
        validator=validate_font_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="指定幻灯片插入",
        fixture_name="simple.pptx",
        description="在 slideIndex=1 (第二张) 上插入文本",
        validator=validate_slide_index_insert,
        tags=["basic"],
    ),
]

_INSERT_PARAMS: list[tuple[str, dict[str, Any] | None]] = [
    ("Hello PPT", None),
    ("定位文本", {"left": 100, "top": 100, "width": 300, "height": 50}),
    ("字体测试", {"font": {"size": 24, "name": "Arial"}}),  # OASP 0.4.0: 嵌套 font（PptFont）
    ("第二页文本", {"slideIndex": 1}),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    text, options = _INSERT_PARAMS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 插入文本 '{text}' (options={options})...")
            start_time = time.time()
            success, data, error = await ppt_insert_text(workspace, fixture.document_uri, text, options=options)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 插入失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}
            print(f"   返回数据: {data}")

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

    parser = argparse.ArgumentParser(description="PPT Insert Text E2E Tests — Basic")
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
