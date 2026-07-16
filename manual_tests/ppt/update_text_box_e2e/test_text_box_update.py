"""
PPT Update TextBox E2E Tests

测试文本框更新功能（有状态工作流: insert → get elementId → update → verify）。

测试场景:
1. 修改文本 — 更新文本内容
2. 字体样式 — 更新 bold/italic
3. 字号字体 — 更新 fontSize/fontName
4. 颜色 — 更新 color/fillColor

运行方式:
    uv run python manual_tests/ppt/update_text_box_e2e/test_text_box_update.py --test all
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
    ppt_insert_text,
    ppt_update_text_box,
)

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def _extract_text_element_id(elements: list[dict[str, Any]]) -> str | None:
    """从元素列表中提取第一个文本类型的 elementId"""
    for elem in elements:
        elem_type = elem.get("type", "").lower()
        if "text" in elem_type or elem_type == "textbox":
            return elem.get("id")
    # 回退: 返回最后一个元素的 id（刚插入的通常在最后）
    if elements:
        return elements[-1].get("id")
    return None


async def _workflow_update_text(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert text → get elementId → update text → verify"""
    # Step 1: 插入文本
    print("\n   📝 Step 1: 插入文本框...")
    success, _, error = await ppt_insert_text(workspace, doc_uri, "原始文本")
    if not success:
        print(f"   ❌ 插入失败: {error}")
        return False
    print("   ✅ 插入成功")

    # Step 2: 获取 elementId
    print("   📝 Step 2: 获取元素列表...")
    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        print(f"   ❌ 获取元素失败: {error}")
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_text_element_id(elements)
    if not element_id:
        print("   ❌ 未找到文本元素 ID")
        return False
    print(f"   ✅ 找到 elementId={element_id}")

    # Step 3: 更新文本
    print("   📝 Step 3: 更新文本内容...")
    success, _, error = await ppt_update_text_box(workspace, doc_uri, element_id, {"text": "更新后的文本"})
    if not success:
        print(f"   ❌ 更新失败: {error}")
        return False
    print("   ✅ 更新成功")
    return True


async def _workflow_update_font_style(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update bold/italic"""
    success, _, error = await ppt_insert_text(workspace, doc_uri, "字体样式测试")
    if not success:
        print(f"   ❌ 插入失败: {error}")
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_text_element_id(elements)
    if not element_id:
        return False

    # OASP 0.4.0: 字体收敛到嵌套 font（PptFont）
    success, _, error = await ppt_update_text_box(workspace, doc_uri, element_id, {"font": {"bold": True, "italic": True}})
    if not success:
        print(f"   ❌ 更新字体样式失败: {error}")
        return False
    print("   ✅ 字体样式更新成功 (bold=True, italic=True)")
    return True


async def _workflow_update_font_size(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update fontSize/fontName"""
    success, _, error = await ppt_insert_text(workspace, doc_uri, "字号字体测试")
    if not success:
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_text_element_id(elements)
    if not element_id:
        return False

    # OASP 0.4.0: 字体收敛到嵌套 font（PptFont）
    success, _, error = await ppt_update_text_box(workspace, doc_uri, element_id, {"font": {"size": 28, "name": "Arial"}})
    if not success:
        print(f"   ❌ 更新字号失败: {error}")
        return False
    print("   ✅ 字号字体更新成功 (fontSize=28, fontName=Arial)")
    return True


async def _workflow_update_color(workspace: Any, doc_uri: str) -> bool:
    """工作流: insert → get → update color/fillColor"""
    success, _, error = await ppt_insert_text(workspace, doc_uri, "颜色测试")
    if not success:
        return False

    success, data, error = await ppt_get_current_slide_elements(workspace, doc_uri)
    if not success:
        return False
    elements = (data or {}).get("elements", [])
    element_id = _extract_text_element_id(elements)
    if not element_id:
        return False

    # OASP 0.4.0: 字体色收敛到嵌套 font（PptFont）；fillColor（文本框填充）仍在顶层
    success, _, error = await ppt_update_text_box(
        workspace, doc_uri, element_id, {"font": {"color": "#FF0000"}, "fillColor": "#FFFF00"}
    )
    if not success:
        print(f"   ❌ 更新颜色失败: {error}")
        return False
    print("   ✅ 颜色更新成功 (color=#FF0000, fillColor=#FFFF00)")
    return True


_WORKFLOW_FUNCS = [
    _workflow_update_text,
    _workflow_update_font_style,
    _workflow_update_font_size,
    _workflow_update_color,
]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="修改文本", fixture_name="empty.pptx", description="插入文本框后修改文本内容", tags=["crud"]),
    PptTestCase(name="字体样式", fixture_name="empty.pptx", description="插入文本框后设置 bold/italic", tags=["crud"]),
    PptTestCase(
        name="字号字体", fixture_name="empty.pptx", description="插入文本框后修改 fontSize/fontName", tags=["crud"]
    ),
    PptTestCase(
        name="颜色", fixture_name="empty.pptx", description="插入文本框后修改 color/fillColor", tags=["crud"]
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

    parser = argparse.ArgumentParser(description="PPT Update TextBox E2E Tests")
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
