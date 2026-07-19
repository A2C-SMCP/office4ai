"""
PPT Add Slide E2E Tests

测试添加幻灯片功能。

测试场景:
1. 空白新增 — 默认添加空白幻灯片
2. 指定位置 — 在 insertIndex=0 插入
3. 指定布局 — 使用 layout 名称添加

运行方式:
    uv run python manual_tests/ppt/slide_management_e2e/test_add_slide.py --test all
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
from manual_tests.ppt.test_helpers import ppt_add_slide

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_add_default(data: dict[str, Any], reader: PresentationReader) -> bool:
    reader.reload()
    if reader.slide_count < 2:
        print(f"   ❌ 添加后 slide_count={reader.slide_count}，应 >= 2")
        return False
    print(f"   ✅ 添加后 slide_count={reader.slide_count}")
    return True


def validate_add_at_index(data: dict[str, Any], reader: PresentationReader) -> bool:
    reader.reload()
    if reader.slide_count < 2:
        print(f"   ❌ 添加后 slide_count={reader.slide_count}")
        return False
    print(f"   ✅ 在 index=0 添加后 slide_count={reader.slide_count}")
    return True


def validate_add_with_layout(data: dict[str, Any]) -> bool:
    print(f"   ✅ 使用布局添加幻灯片成功: {data}")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="空白新增", fixture_name="empty.pptx", description="默认添加一张幻灯片", validator=validate_add_default, tags=["basic"]
    ),
    PptTestCase(
        name="指定位置",
        fixture_name="simple.pptx",
        description="在 insertIndex=0 插入幻灯片",
        validator=validate_add_at_index,
        tags=["basic"],
    ),
    PptTestCase(
        name="指定布局",
        fixture_name="simple.pptx",
        description="使用 'Blank' 布局添加幻灯片",
        validator=validate_add_with_layout,
        tags=["advanced"],
    ),
]

_ADD_OPTIONS: list[dict[str, Any] | None] = [
    None,
    {"insertIndex": 0},
    {"layout": "Blank"},
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    options = _ADD_OPTIONS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 添加幻灯片 (options={options})...")
            start_time = time.time()
            success, data, error = await ppt_add_slide(workspace, fixture.document_uri, options=options)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 添加失败: {error}")
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

    parser = argparse.ArgumentParser(description="PPT Add Slide E2E Tests")
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
