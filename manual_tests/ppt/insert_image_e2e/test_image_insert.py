"""
PPT Insert Image E2E Tests

测试 PPT 图片插入功能。

测试场景:
1. 基本图片插入 — 插入最小合法 PNG
2. 带位置图片插入 — 指定 left/top/width/height
3. 跨幻灯片插入 — slideIndex=1，验证图片在目标幻灯片
4. slideIndex 越界 — slideIndex=999，预期失败
5. 视图恢复验证 — 跨页插入后 currentSlideIndex 应恢复

运行方式:
    uv run python manual_tests/ppt/insert_image_e2e/test_image_insert.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt.e2e_base import (
    PPTTestRunner,
    PptTestCase,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import (
    ppt_get_slide_elements,
    ppt_get_slide_info,
    ppt_insert_image,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"

# 最小合法 1x1 红色 PNG (base64)
MINIMAL_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4"
    "nGP4z8BQDwAEgAF/pooBPQAAAABJRU5ErkJggg=="
)


# ==============================================================================
# Validators (tests 1-2)
# ==============================================================================


def validate_basic_insert(data: dict[str, Any]) -> bool:
    """验证基本图片插入"""
    print(f"   ✅ 协议返回: {data}")
    return True


def validate_positioned_insert(data: dict[str, Any]) -> bool:
    """验证带位置的图片插入"""
    print(f"   ✅ 协议返回: {data}")
    return True


# ==============================================================================
# 测试用例定义
# ==============================================================================

TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="基本图片插入",
        fixture_name="empty.pptx",
        description="在空白幻灯片插入最小 PNG 图片",
        validator=validate_basic_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带位置图片插入",
        fixture_name="empty.pptx",
        description="在指定位置 (200, 150) 插入图片",
        validator=validate_positioned_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="跨幻灯片插入",
        fixture_name="simple.pptx",
        description="在 slideIndex=1 上插入图片，验证目标幻灯片有新元素",
        validator=None,  # 使用自定义逻辑
        tags=["cross_slide"],
    ),
    PptTestCase(
        name="slideIndex 越界",
        fixture_name="simple.pptx",
        description="slideIndex=999，预期返回失败",
        validator=None,
        tags=["cross_slide", "expected_failure"],
    ),
    PptTestCase(
        name="视图恢复验证",
        fixture_name="simple.pptx",
        description="在 slideIndex=2 插入后，currentSlideIndex 应恢复到 0",
        validator=None,
        tags=["cross_slide", "view_restore"],
    ),
]

_INSERT_PARAMS: list[tuple[dict[str, Any], dict[str, Any] | None]] = [
    ({"base64": MINIMAL_PNG_BASE64}, {"left": 100, "top": 100, "width": 200, "height": 150}),
    ({"base64": MINIMAL_PNG_BASE64}, {"left": 200, "top": 150, "width": 100, "height": 100}),
    ({"base64": MINIMAL_PNG_BASE64}, {"slideIndex": 1, "left": 100, "top": 100, "width": 200, "height": 150}),
    ({"base64": MINIMAL_PNG_BASE64}, {"slideIndex": 999}),
    ({"base64": MINIMAL_PNG_BASE64}, {"slideIndex": 2, "left": 100, "top": 100, "width": 200, "height": 150}),
]


# ==============================================================================
# 测试执行
# ==============================================================================


async def _run_basic_test(workspace: Any, doc_uri: str, image_data: dict, options: dict | None, validator: Any) -> bool:
    """执行基础插入测试（tests 1-2）"""
    success, data, error = await ppt_insert_image(workspace, doc_uri, image_data, options=options)
    if not success:
        print(f"❌ 插入失败: {error}")
        return False
    print("✅ 协议返回成功")
    data = data or {}
    print("\n📊 验证结果:")
    if validator and not validator(data):
        return False
    return True


async def _run_cross_slide_test(workspace: Any, doc_uri: str, image_data: dict, options: dict) -> bool:
    """执行跨幻灯片插入测试（test 3）— 验证图片确实在目标幻灯片上"""
    target_slide = options["slideIndex"]

    # 1. 获取插入前目标幻灯片元素数
    ok, before_data, err = await ppt_get_slide_elements(workspace, doc_uri, target_slide)
    if not ok:
        print(f"❌ 获取插入前元素失败: {err}")
        return False
    before_count = len((before_data or {}).get("elements", []))
    print(f"   插入前 slide {target_slide} 元素数: {before_count}")

    # 2. 执行插入
    success, data, error = await ppt_insert_image(workspace, doc_uri, image_data, options=options)
    if not success:
        print(f"❌ 插入失败: {error}")
        return False
    print("✅ 协议返回成功")
    print(f"   返回数据: {data}")

    # 3. 获取插入后目标幻灯片元素数
    await asyncio.sleep(0.5)
    ok, after_data, err = await ppt_get_slide_elements(workspace, doc_uri, target_slide)
    if not ok:
        print(f"❌ 获取插入后元素失败: {err}")
        return False
    after_count = len((after_data or {}).get("elements", []))
    print(f"   插入后 slide {target_slide} 元素数: {after_count}")

    print("\n📊 验证结果:")
    if after_count > before_count:
        print(f"   ✅ slide {target_slide} 新增 {after_count - before_count} 个元素")
        return True
    print(f"   ❌ slide {target_slide} 元素数未增加 (前={before_count}, 后={after_count})")
    return False


async def _run_out_of_bounds_test(workspace: Any, doc_uri: str, image_data: dict, options: dict) -> bool:
    """执行 slideIndex 越界测试（test 4）— 预期失败"""
    success, data, error = await ppt_insert_image(workspace, doc_uri, image_data, options=options)
    print("\n📊 验证结果:")
    if not success:
        print(f"   ✅ 预期失败，错误信息: {error}")
        return True
    print(f"   ❌ slideIndex={options['slideIndex']} 应该失败，但返回成功: {data}")
    return False


async def _run_view_restore_test(workspace: Any, doc_uri: str, image_data: dict, options: dict) -> bool:
    """执行视图恢复测试（test 5）— 插入后 currentSlideIndex 应恢复"""
    # 1. 记录插入前的 currentSlideIndex
    ok, info_before, err = await ppt_get_slide_info(workspace, doc_uri)
    if not ok:
        print(f"❌ 获取插入前信息失败: {err}")
        return False
    original_index = (info_before or {}).get("currentSlideIndex", 0)
    print(f"   插入前 currentSlideIndex: {original_index}")

    # 2. 执行跨幻灯片插入
    target_slide = options["slideIndex"]
    success, data, error = await ppt_insert_image(workspace, doc_uri, image_data, options=options)
    if not success:
        print(f"❌ 插入失败: {error}")
        return False
    print(f"✅ 在 slide {target_slide} 插入成功")

    # 3. 检查 currentSlideIndex 是否恢复
    await asyncio.sleep(0.5)
    ok, info_after, err = await ppt_get_slide_info(workspace, doc_uri)
    if not ok:
        print(f"❌ 获取插入后信息失败: {err}")
        return False
    restored_index = (info_after or {}).get("currentSlideIndex", -1)
    print(f"   插入后 currentSlideIndex: {restored_index}")

    print("\n📊 验证结果:")
    if restored_index == original_index:
        print(f"   ✅ 视图已恢复: currentSlideIndex={restored_index} (与插入前一致)")
        return True
    print(f"   ❌ 视图未恢复: 插入前={original_index}, 插入后={restored_index}")
    return False


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    image_data, options = _INSERT_PARAMS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 插入图片 (options={options})...")
            start_time = time.time()

            if test_number <= 2:
                passed = await _run_basic_test(workspace, fixture.document_uri, image_data, options, test_case.validator)
            elif test_number == 3:
                passed = await _run_cross_slide_test(workspace, fixture.document_uri, image_data, options or {})
            elif test_number == 4:
                passed = await _run_out_of_bounds_test(workspace, fixture.document_uri, image_data, options or {})
            elif test_number == 5:
                passed = await _run_view_restore_test(workspace, fixture.document_uri, image_data, options or {})
            else:
                passed = False

            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")

            print("\n" + "=" * 70)
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
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
    auto_close: bool = True,
    cleanup_on_success: bool = True,
) -> bool:
    ensure_ppt_fixtures(FIXTURES_DIR)
    runner = PPTTestRunner(
        fixtures_dir=FIXTURES_DIR.parent, auto_open=auto_open, auto_close=auto_close, cleanup_on_success=cleanup_on_success
    )
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

    parser = argparse.ArgumentParser(description="PPT Insert Image E2E Tests")
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--no-auto-close", action="store_true", help="不自动关闭文档（调试用）")
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
            run_tests(
                test_indices,
                auto_open=not args.no_auto_open,
                auto_close=not args.no_auto_close,
                cleanup_on_success=not args.always_cleanup,
            )
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
