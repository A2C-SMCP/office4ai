"""
PPT Update Image E2E Tests

测试图片更新功能（有状态工作流: insert image → get elementId → update image）。

测试场景:
1. 替换图片 — 用新图片替换已有图片
2. keepDimensions — 替换图片时保持原尺寸

运行方式:
    uv run python manual_tests/ppt/update_image_e2e/test_image_update.py --test all
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
    ppt_get_current_slide_elements,
    ppt_insert_image,
    ppt_update_image,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"

# 最小合法 1x1 PNG (red/blue)
PNG_RED = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4"
    "nGP4z8BQDwAEgAF/pooBPQAAAABJRU5ErkJggg=="
)
PNG_BLUE = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADklEQVQI"
    "12P4z8BQDwAEgAF/pooBPQAAAABJRU5ErkJggg=="
)


def _extract_image_element_id(elements: list[dict[str, Any]]) -> str | None:
    """提取第一个图片类型的 elementId"""
    for elem in elements:
        elem_type = elem.get("type", "").lower()
        if "image" in elem_type or "picture" in elem_type:
            return elem.get("id")
    if elements:
        return elements[-1].get("id")
    return None


async def _workflow_replace_image(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update image"""
    print("\n   📝 Step 1: 插入图片...")
    success, _, error = await ppt_insert_image(workspace, doc_uri, {"base64": PNG_RED})
    if not success:
        print(f"   ❌ 插入失败: {error}")
        return False

    print("   📝 Step 2: 获取元素...")
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_image_element_id(elements)
    if not element_id:
        print("   ❌ 未找到图片元素 ID")
        return False
    print(f"   ✅ elementId={element_id}")

    print("   📝 Step 3: 替换图片...")
    success, _, error = await ppt_update_image(workspace, doc_uri, element_id, {"base64": PNG_BLUE})
    if not success:
        print(f"   ❌ 替换失败: {error}")
        return False
    print("   ✅ 图片替换成功")
    return True


async def _workflow_keep_dimensions(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update with keepDimensions"""
    success, _, error = await ppt_insert_image(
        workspace, doc_uri, {"base64": PNG_RED}, options={"width": 200, "height": 150}
    )
    if not success:
        print(f"   ❌ 插入失败: {error}")
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_image_element_id(elements)
    if not element_id:
        return False

    success, _, error = await ppt_update_image(
        workspace, doc_uri, element_id, {"base64": PNG_BLUE}, options={"keepDimensions": True}
    )
    if not success:
        print(f"   ❌ keepDimensions 替换失败: {error}")
        return False
    print("   ✅ keepDimensions 替换成功")
    return True


_WORKFLOW_FUNCS = [_workflow_replace_image, _workflow_keep_dimensions]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="替换图片", fixture_name="empty.pptx", description="插入红色 PNG 后替换为蓝色 PNG", tags=["crud"]),
    PptTestCase(
        name="keepDimensions",
        fixture_name="empty.pptx",
        description="替换图片时保持原始尺寸",
        tags=["crud"],
    ),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    workflow_func = _WORKFLOW_FUNCS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            passed = await workflow_func(workspace, fixture.document_uri)
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

    parser = argparse.ArgumentParser(description="PPT Update Image E2E Tests")
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
