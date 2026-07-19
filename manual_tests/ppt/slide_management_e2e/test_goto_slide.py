"""
PPT Goto Slide E2E Tests

测试幻灯片跳转功能。

测试场景:
1. 跳到首页 — slideIndex=0
2. 跳到末页 — slideIndex=最后一张
3. 错误码 3008 — slideIndex 越界（oasp#23，Add-In 接线前为 XFAIL）

运行方式:
    uv run python manual_tests/ppt/slide_management_e2e/test_goto_slide.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.error_case import PENDING_ADDIN_92, judge_error_case
from manual_tests.ppt.e2e_base import (
    PptTestCase,
    PPTTestRunner,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_get_slide_info, ppt_goto_slide

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


async def _workflow_goto_first(workspace: Any, doc_uri: str) -> bool:
    """跳到首页"""
    success, _, error = await ppt_goto_slide(workspace, doc_uri, slide_index=0)
    if not success:
        print(f"   ❌ 跳转失败: {error}")
        return False

    # 验证 currentSlideIndex
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    current = (data or {}).get("currentSlideIndex", -1)
    if current != 0:
        print(f"   ⚠️  currentSlideIndex={current}，预期 0 (部分 Add-In 可能不返回此字段)")
    print(f"   ✅ 跳转到首页成功 (currentSlideIndex={current})")
    return True


async def _workflow_goto_last(workspace: Any, doc_uri: str) -> bool:
    """跳到末页"""
    # 先获取 slideCount
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    count = (data or {}).get("slideCount", 0)
    if count < 2:
        print(f"   ❌ 需要 >= 2 张幻灯片")
        return False
    last_index = count - 1

    success, _, error = await ppt_goto_slide(workspace, doc_uri, slide_index=last_index)
    if not success:
        print(f"   ❌ 跳转失败: {error}")
        return False

    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    current = (data or {}).get("currentSlideIndex", -1)
    print(f"   ✅ 跳转到末页成功 (slideIndex={last_index}, currentSlideIndex={current})")
    return True


async def _workflow_goto_out_of_range(workspace: Any, doc_uri: str) -> bool:
    """slideIndex 越界 → 3008 POSITION_INVALID + details（oasp#23 / issue #90）。

    与 ``ppt:delete:slide`` 同判法。oasp#23 仅取这两个 /ppt 事件作样板验证判法，
    全量清扫见 oasp#24——**其余序号越界事件沿用其表内现码，勿照抄本用例扩大化**。
    """
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    total = (data or {}).get("slideCount", 0)
    out_of_range = total + 100

    print(f"\n   📋 slideCount={total}，请求跳转 index={out_of_range}（越界）")
    ok, _, error = await ppt_goto_slide(workspace, doc_uri, slide_index=out_of_range)

    verdict, message = judge_error_case(
        ok,
        error,
        expect_code="3008",
        expect_details={"kind": "slide"},
        xfail_reason=PENDING_ADDIN_92,
    )
    print(message)
    return verdict.passed


_WORKFLOW_FUNCS = [_workflow_goto_first, _workflow_goto_last, _workflow_goto_out_of_range]

TEST_CASES: list[PptTestCase] = [
    PptTestCase(name="跳到首页", fixture_name="multi_slide.pptx", description="跳转到 slideIndex=0", tags=["basic"]),
    PptTestCase(
        name="跳到末页", fixture_name="multi_slide.pptx", description="跳转到最后一张幻灯片", tags=["basic"]
    ),
    PptTestCase(
        name="错误码 3008 — slideIndex 越界",
        fixture_name="multi_slide.pptx",
        description="goto:slide 请求越界序号 → 3008 POSITION_INVALID（oasp#23，过渡期样板事件）",
        expect_error_code="3008",
        expect_error_details={"kind": "slide"},
        xfail_reason=PENDING_ADDIN_92,
        tags=["error"],
    ),
]
if len(_WORKFLOW_FUNCS) != len(TEST_CASES):  # 结构约束，非调试断言（python -O 会剥离 assert）
    raise RuntimeError(f"TEST_CASES({len(TEST_CASES)}) 与 _WORKFLOW_FUNCS({len(_WORKFLOW_FUNCS)}) 必须一一对应")


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            passed = await _WORKFLOW_FUNCS[test_number - 1](workspace, fixture.document_uri)
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  总执行时间: {elapsed_ms:.1f}ms")
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
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

    parser = argparse.ArgumentParser(description="PPT Goto Slide E2E Tests")
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
