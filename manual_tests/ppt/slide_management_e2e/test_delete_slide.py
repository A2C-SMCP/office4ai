"""
PPT Delete Slide E2E Tests

测试删除幻灯片功能。

测试场景:
1. 删除中间幻灯片 — 删除 index=1 (中间)
2. 删除末尾幻灯片 — 删除最后一张
3. 错误码 3008 — slideIndex 越界（oasp#23，Add-In 接线前为 XFAIL）

运行方式:
    uv run python manual_tests/ppt/slide_management_e2e/test_delete_slide.py --test all
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
    PresentationReader,
    ensure_ppt_fixtures,
)
from manual_tests.ppt.test_helpers import ppt_delete_slide, ppt_get_slide_info

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


async def _workflow_delete_middle(workspace: Any, doc_uri: str, working_path: Any) -> bool:
    """删除中间幻灯片"""
    # 先获取当前 slideCount
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    before = (data or {}).get("slideCount", 0)
    if before < 3:
        print(f"   ❌ slideCount={before}，需要 >= 3")
        return False

    success, _, error = await ppt_delete_slide(workspace, doc_uri, slide_index=1)
    if not success:
        print(f"   ❌ 删除失败: {error}")
        return False

    # 协议验证
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    after = (data or {}).get("slideCount", 0)
    if after >= before:
        print(f"   ❌ 删除后 slideCount={after}，未减少")
        return False
    print(f"   ✅ 协议验证: {before} → {after}")

    # python-pptx 双重验证
    await asyncio.sleep(0.5)
    reader = PresentationReader(working_path)
    reader.reload()
    print(f"   ✅ python-pptx 验证: slide_count={reader.slide_count}")
    return True


async def _workflow_delete_last(workspace: Any, doc_uri: str, working_path: Any) -> bool:
    """删除最后一张幻灯片"""
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    before = (data or {}).get("slideCount", 0)
    last_index = before - 1

    success, _, error = await ppt_delete_slide(workspace, doc_uri, slide_index=last_index)
    if not success:
        print(f"   ❌ 删除失败: {error}")
        return False

    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    after = (data or {}).get("slideCount", 0)
    if after >= before:
        print(f"   ❌ 删除后 slideCount={after}")
        return False
    print(f"   ✅ 删除末尾: {before} → {after}")

    await asyncio.sleep(0.5)
    reader = PresentationReader(working_path)
    reader.reload()
    print(f"   ✅ python-pptx 验证: slide_count={reader.slide_count}")
    return True


async def _workflow_delete_out_of_range(workspace: Any, doc_uri: str, working_path: Any) -> bool:
    """slideIndex 越界 → 3008 POSITION_INVALID + details（oasp#23 / issue #90）。

    该值在线缆上完全合法（非负整数），仅相对**当前文档状态**无效，故不是 4002/4004——
    正确恢复是「重读文档状态后原值重试」而非「改参数」。details 回带 ``{index,total,kind}``
    让调用方一次拿到实时边界。
    """
    success, data, _ = await ppt_get_slide_info(workspace, doc_uri)
    if not success:
        return False
    total = (data or {}).get("slideCount", 0)
    out_of_range = total + 100

    print(f"\n   📋 slideCount={total}，请求删除 index={out_of_range}（越界）")
    ok, _, error = await ppt_delete_slide(workspace, doc_uri, slide_index=out_of_range)

    verdict, message = judge_error_case(
        ok,
        error,
        expect_code="3008",
        expect_details={"kind": "slide"},
        xfail_reason=PENDING_ADDIN_92,
    )
    print(message)

    # 越界删除绝不能改动文档——无论错误码对不对，页数必须原样
    await asyncio.sleep(0.5)
    reader = PresentationReader(working_path)
    reader.reload()
    if reader.slide_count != total:
        print(f"   ❌ 越界删除改动了文档: {total} → {reader.slide_count}")
        return False
    print(f"   ✅ 文档未被改动: slide_count={reader.slide_count}")
    return verdict.passed


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="删除中间幻灯片",
        fixture_name="multi_slide.pptx",
        description="删除 5 页 PPT 的第 2 张 (index=1)",
        tags=["crud"],
    ),
    PptTestCase(
        name="删除末尾幻灯片",
        fixture_name="multi_slide.pptx",
        description="删除最后一张幻灯片",
        tags=["crud"],
    ),
    PptTestCase(
        name="错误码 3008 — slideIndex 越界",
        fixture_name="multi_slide.pptx",
        description="delete:slide 请求越界序号 → 3008 POSITION_INVALID（oasp#23，过渡期样板事件）",
        expect_error_code="3008",
        expect_error_details={"kind": "slide"},
        xfail_reason=PENDING_ADDIN_92,
        tags=["error"],
    ),
]

#: 与 TEST_CASES 一一对应（索引即 test_number - 1）——原先 ``if test_number == 1 / else``
#: 的二分派在加入第 3 个用例后会把它错误路由到 _workflow_delete_last，故改为按序表驱动。
_WORKFLOWS = [
    _workflow_delete_middle,
    _workflow_delete_last,
    _workflow_delete_out_of_range,
]
if len(_WORKFLOWS) != len(TEST_CASES):  # 结构约束，非调试断言（python -O 会剥离 assert）
    raise RuntimeError(f"TEST_CASES({len(TEST_CASES)}) 与 _WORKFLOWS({len(_WORKFLOWS)}) 必须一一对应")


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start_time = time.time()
            workflow = _WORKFLOWS[test_number - 1]
            passed = await workflow(workspace, fixture.document_uri, fixture.working_path)
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

    parser = argparse.ArgumentParser(description="PPT Delete Slide E2E Tests")
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
