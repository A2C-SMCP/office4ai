"""
PPT Insert Shape E2E Tests — Basic

测试基本的 PPT 形状插入功能。

测试场景:
1. 矩形 — 插入 Rectangle
2. 圆形 — 插入 Circle
3. 带文本 — 插入带 text 选项的形状
4. 带样式 — 插入带 fillColor/borderColor 的形状
5. 带字体（OASP 0.4.0）— text-capable 形状 insert-with-font，读回 python-pptx 校验 run 字体
6. Line 拒绝 font（负例）— 无文本框的 Line 传 font/text → 期望 4002 静态拒绝

运行方式:
    uv run python manual_tests/ppt/insert_shape_e2e/test_shape_insert.py --test all
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
from manual_tests.ppt.test_helpers import ppt_insert_shape

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_shape_insert(data: dict[str, Any]) -> bool:
    """通用形状插入验证（仅协议返回）"""
    print(f"   ✅ 协议返回: {data}")
    return True


def validate_shape_font(data: dict[str, Any], reader: PresentationReader) -> bool:
    """双重验证（OASP 0.4.0 insert-with-font）：读回 pptx 校验形状 run 字体是否真的生效。

    校验点（office.js → PowerPoint → pptx 往返后应持久化）：
      - bold=True                     → <a:rPr b="1">
      - size≈24pt                     → <a:rPr sz="2400">
      - doubleStrikethrough           → <a:rPr strike="dblStrike">（camelCase wire 属性的落地证据）
    字体名（微软雅黑）作 informational——宿主字体替换不阻塞。
    """
    reader.reload()

    target = None
    for slide in reader.prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame or "标题" not in shape.text_frame.text:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        target = run
                        break
                if target is not None:
                    break
            if target is not None:
                break
        if target is not None:
            break

    if target is None:
        print("   ❌ 未找到含 '标题' 的带文本形状（insert-with-font 可能未落地）")
        return False

    font = target.font
    size_pt = font.size.pt if font.size is not None else None
    strike = font._rPr.get("strike")
    try:
        color = str(font.color.rgb) if font.color is not None and font.color.type is not None else None
    except Exception:
        color = None

    print(
        f"   🔎 读回 run={target.text!r} → bold={font.bold} size={size_pt} "
        f"name={font.name!r} color={color} strike={strike}"
    )

    ok = True
    if font.bold is not True:
        print("   ❌ bold 未生效（期望 True）")
        ok = False
    if size_pt is None or abs(size_pt - 24) > 0.6:
        print(f"   ❌ size 未生效（期望 24pt，实际 {size_pt}）")
        ok = False
    if strike != "dblStrike":
        print(f"   ❌ doubleStrikethrough 未生效（期望 strike='dblStrike'，实际 {strike!r}）")
        ok = False
    if font.name and font.name != "微软雅黑":
        print(f"   ⚠️  字体名 {font.name!r} ≠ '微软雅黑'（可能宿主字体替换，非阻塞）")
    if ok:
        print("   ✅ 文档内容验证通过: bold + size24 + doubleStrikethrough 均已生效")
    return ok


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="插入矩形",
        fixture_name="empty.pptx",
        description="插入 Rectangle 形状",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="插入圆形",
        fixture_name="empty.pptx",
        description="插入 Circle 形状",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带文本形状",
        fixture_name="empty.pptx",
        description="插入带 text='Hello Shape' 的矩形",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带样式形状",
        fixture_name="empty.pptx",
        description="插入带红色填充和蓝色边框的矩形",
        validator=validate_shape_insert,
        tags=["advanced"],
    ),
    PptTestCase(
        name="带字体形状 (0.4.0)",
        fixture_name="empty.pptx",
        description="插入 RoundedRectangle 带 text+font(size24/微软雅黑/红/粗/双删除线)，读回校验字体",
        validator=validate_shape_font,
        tags=["font", "0.4.0"],
    ),
    PptTestCase(
        name="Line 拒绝 font (负例, 0.4.0)",
        fixture_name="empty.pptx",
        description="无文本框的 Line 传 font/text → 期望 4002 静态拒绝（text-capable 门控）",
        validator=None,
        tags=["font", "0.4.0", "negative"],
    ),
]

# (shape_type, options, expect)：expect=None 期望成功；expect={"fail_error_contains": "4002"} 期望失败
_SHAPE_PARAMS: list[tuple[str, dict[str, Any] | None, dict[str, str] | None]] = [
    ("Rectangle", {"left": 100, "top": 100, "width": 200, "height": 150}, None),
    ("Circle", {"left": 300, "top": 100, "width": 150, "height": 150}, None),
    ("Rectangle", {"text": "Hello Shape", "left": 100, "top": 300, "width": 250, "height": 100}, None),
    (
        "Rectangle",
        {
            "fillColor": "#FF0000",
            "borderColor": "#0000FF",
            "borderWidth": 3,
            "left": 100,
            "top": 100,
            "width": 200,
            "height": 150,
        },
        None,
    ),
    (
        "RoundedRectangle",
        {
            "text": "标题",
            "left": 100,
            "top": 100,
            "width": 320,
            "height": 120,
            # snake_case 入参（DTO populate_by_name 接受）→ wire 出 camelCase
            "font": {
                "size": 24,
                "name": "微软雅黑",
                "color": "#C00000",
                "bold": True,
                "double_strikethrough": True,
            },
        },
        None,
    ),
    (
        "Line",
        {
            "text": "on a line",
            "left": 100,
            "top": 300,
            "width": 300,
            "height": 10,
            "font": {"size": 18, "bold": True},
        },
        {"fail_error_contains": "4002"},
    ),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    shape_type, options, expect = _SHAPE_PARAMS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 插入形状 '{shape_type}' (options={options})...")
            start_time = time.time()
            success, data, error = await ppt_insert_shape(workspace, fixture.document_uri, shape_type, options=options)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            # 负例：期望被 Add-In 前置静态拒绝（text-capable 门控）
            if expect is not None:
                want = expect["fail_error_contains"]
                if success:
                    print(f"❌ 期望失败（错误含 '{want}'），但插入却成功了 —— text-capable 门控未生效")
                    return False
                if want not in (error or ""):
                    print(f"❌ 期望错误含 '{want}'，实际错误: {error!r}")
                    return False
                print(f"✅ 如期被静态拒绝（含 '{want}'）: {error}")
                print("\n" + "=" * 70)
                print(f"✅ 测试 {test_number} 通过（负例）")
                print("=" * 70)
                return True

            # 正例：期望成功
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

    parser = argparse.ArgumentParser(description="PPT Insert Shape E2E Tests — Basic")
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
            run_tests(test_indices, auto_open=not args.no_auto_open, cleanup_on_success=not args.always_cleanup or True)
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
