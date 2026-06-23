"""
Word OOXML Round-Trip E2E Tests (自动化版本)

测试 word_get_ooxml / word_insert_ooxml 工具对的端到端 round-trip。

⚠️ 路径收窄(#46-D):word_insert_ooxml 仅吃**片段**;整篇文档包会被 Office.js insertOoxml
   以 GeneralException 拒绝。整篇 round-trip 已移到 document_file_e2e(base64 .docx 路)。
   故本套只覆盖**片段/选区**语义。

测试场景:
1. 选区片段 round-trip — get(scope=selection)→落盘片段 Flat OPC→insert(Replace, selection)→
   协议成功(需先在 Word 手动选中一段文字;空选区会回退整篇,本场景将提示改用 document_file)
2. Body OOXML 只读导出 — get(scope=body) 落盘,校验 Flat OPC、data 不回 inline ooxml、
   scope 回生效值(整篇包仅可读,不经 word_insert_ooxml 回灌)

设计要点:
- 驱动 MCP 工具(WordGetOoxmlTool / WordInsertOoxmlTool)而非裸 OfficeAction，
  因为 file-handle(dest_path/source_path)读写逻辑在工具 execute 内。
- Word OOXML(片段)是 Flat OPC 字符串，落盘/读取均为文本，无 base64。

⚠️ 外部依赖(真机): 需 office-editor4ai (word-editor4ai) Add-In 已实现
   word:get:ooxml / word:insert:ooxml handler(cross-ask 已锁定契约，见 office4ai#46)。

运行方式:
    uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --test 1
    uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --test all
    uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --list
"""

import asyncio
import sys
import time
from pathlib import Path

from manual_tests.e2e_base import (
    DocumentFixture,
    DocumentReader,
    E2ETestRunner,
    TestCase,
    ensure_fixtures,
)
from office4ai.a2c_smcp.tools.word import WordGetOoxmlTool, WordInsertOoxmlTool
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 配置
# ==============================================================================

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ooxml_e2e"

# simple.docx 夹具中已知存在的文本(ensure_fixtures 生成)，用于 round-trip 保真校验
KNOWN_TEXT_CANDIDATES = ["测试", "文本", "Hello", "段落"]

# Flat OPC 包根标记(range.getOoxml() 产出的 WordprocessingML Flat OPC)
FLAT_OPC_MARKERS = ["<pkg:package", "<?xml", "wordprocessingml", "<w:document", "<w:body"]


# ==============================================================================
# 校验辅助
# ==============================================================================


def _looks_like_flat_opc(ooxml: str) -> bool:
    """粗判一段字符串是否为 Word Flat OPC / WordprocessingML(命中任一标记)。"""
    lower = ooxml.lower()
    return any(m.lower() in lower for m in FLAT_OPC_MARKERS)


def _known_text_in(reader: DocumentReader) -> str | None:
    """返回文档中命中的第一个已知文本，用于保真断言。"""
    for candidate in KNOWN_TEXT_CANDIDATES:
        if reader.contains(candidate):
            return candidate
    return None


# ==============================================================================
# 测试用例
# ==============================================================================

TEST_CASES: list[TestCase] = [
    TestCase(
        name="选区片段 round-trip",
        fixture_name="simple.docx",
        description="get(selection)→落盘片段 Flat OPC→insert(Replace, selection)(需先手动选中文字;空选区提示改用 document_file)",
        tags=["roundtrip"],
    ),
    TestCase(
        name="Body OOXML 只读导出",
        fixture_name="simple.docx",
        description="get(body) 落盘，校验 Flat OPC、data 不回 inline ooxml、scope 回生效值(整篇仅可读,不回灌)",
        tags=["export"],
    ),
]


# ==============================================================================
# 测试执行
# ==============================================================================


async def _run_selection_fragment_roundtrip(workspace: OfficeWorkspace, fixture: DocumentFixture) -> bool:
    """场景 1: 选区片段 round-trip(需先在 Word 手动选中一段文字)。

    收窄说明(#46-D):word_insert_ooxml 仅吃片段;整篇 body 包会被 Office.js insertOoxml 以
    GeneralException 拒绝。整篇 round-trip 请走 document_file_e2e(base64 .docx 路)。
    """
    get_tool = WordGetOoxmlTool(workspace)
    insert_tool = WordInsertOoxmlTool(workspace)
    dest = fixture.working_path.parent / "exported_selection_ooxml.xml"

    # 1) 导出当前选区的 OOXML 片段
    print("\n📝 步骤 1: word_get_ooxml(scope=selection) → 落盘片段 Flat OPC...")
    get_result = await get_tool.execute(
        {"document_uri": fixture.document_uri, "scope": "selection", "dest_path": str(dest)}
    )
    if not get_result.get("success"):
        print(f"   ❌ get 失败: {get_result.get('error')}")
        return False
    effective = get_result["data"].get("scope")
    if effective != "selection":
        # 空选区 → 回退整篇 body,整篇包不能经 word_insert_ooxml 回灌
        print(f"   ⚠️ 生效 scope={effective}(无选区,已回退整篇)。")
        print("      请在 Word 中手动选中一段文字后重试本场景;整篇 round-trip 请用 document_file_e2e。")
        return False
    ooxml = dest.read_text(encoding="utf-8")
    if not _looks_like_flat_opc(ooxml):
        print(f"   ❌ 落盘内容不像 Flat OPC，前 120 字符: {ooxml[:120]}")
        return False
    print(f"   ✅ 导出选区片段 Flat OPC {len(ooxml)} 字符 → {dest.name}")

    # 2) 把片段原地 Replace 回选区(片段 round-trip)
    print("\n📝 步骤 2: word_insert_ooxml(source_path, insertLocation=Replace, scope=selection)...")
    insert_result = await insert_tool.execute(
        {
            "document_uri": fixture.document_uri,
            "source_path": str(dest),
            "insertLocation": "Replace",
            "scope": "selection",
        }
    )
    if not insert_result.get("success"):
        print(f"   ❌ insert 失败: {insert_result.get('error')}")
        return False
    print("   ✅ insert 协议返回成功(片段原地回灌)")

    # 3) 双重验证: 读回文档,已知文本应仍在(片段 round-trip 未丢内容)
    print("\n📊 步骤 3: 文档内容保真校验(DocumentReader)...")
    reader = DocumentReader(fixture.working_path)
    reader.reload()  # 强制 Word 落盘后重读
    hit = _known_text_in(reader)
    if hit is None:
        print(f"   ❌ round-trip 后未找到任何已知文本 {KNOWN_TEXT_CANDIDATES}")
        print(f"      文档前 200 字符: {reader.text[:200]}")
        return False
    print(f"   ✅ round-trip 后已知文本仍在: '{hit}'")
    return True


async def _run_body_export(workspace: OfficeWorkspace, fixture: DocumentFixture) -> bool:
    """场景 2: Body OOXML 导出(仅 get)。"""
    get_tool = WordGetOoxmlTool(workspace)
    dest = fixture.working_path.parent / "exported_only_ooxml.xml"

    print("\n📝 执行: word_get_ooxml(scope=body) → 落盘...")
    result = await get_tool.execute({"document_uri": fixture.document_uri, "scope": "body", "dest_path": str(dest)})
    if not result.get("success"):
        print(f"   ❌ get 失败: {result.get('error')}")
        return False

    data = result.get("data", {})
    passed = True
    # data 不回 inline ooxml
    if "ooxml" in data:
        print("   ❌ data 不应回传 inline ooxml")
        passed = False
    else:
        print("   ✅ data 未回传 inline ooxml(留在磁盘)")
    # scope 回生效值
    if data.get("scope") != "body":
        print(f"   ❌ data.scope 非 body: {data.get('scope')}")
        passed = False
    else:
        print("   ✅ data.scope == body")
    # 落盘为非空 Flat OPC
    if not dest.is_file() or not _looks_like_flat_opc(dest.read_text(encoding="utf-8")):
        print("   ❌ dest_path 未落盘或非 Flat OPC")
        passed = False
    else:
        print(f"   ✅ dest_path 落盘 Flat OPC，bytes={data.get('bytes')}")
    return passed


# 每个测试用例对应的执行函数
_RUNNERS = [_run_selection_fragment_roundtrip, _run_body_export]


async def run_single_test(runner: E2ETestRunner, test_case: TestCase, test_number: int) -> bool:
    """执行单个测试用例。"""
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"ooxml_e2e/{test_case.fixture_name}"
    scenario = _RUNNERS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            start = time.time()
            passed = await scenario(workspace, fixture)
            print(f"\n⏱️  执行时间: {(time.time() - start) * 1000:.1f}ms")

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
    cleanup_on_success: bool = True,
) -> bool:
    """运行指定的测试。"""
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
        if len(test_indices) > 1 and results:
            print("\n" + "-" * 70)
            print("⏳ 准备下一个测试...")
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续...")
        results.append(await run_single_test(runner, TEST_CASES[idx - 1], idx))

    if len(results) > 1:
        print("\n" + "=" * 70)
        print(f"📈 总体结果: {sum(results)}/{len(results)} 测试通过")
        print("=" * 70)
    return all(results) if results else False


# ==============================================================================
# 命令行入口
# ==============================================================================


def main() -> None:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Word OOXML Round-Trip E2E Tests (自动化版本)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=选区片段 round-trip, 2=Body 只读导出, all=全部",
    )
    parser.add_argument("--no-auto-open", action="store_true", help="不自动打开文档")
    parser.add_argument("--always-cleanup", action="store_true", help="无论成功失败都清理")
    parser.add_argument("--list", action="store_true", help="列出所有测试用例")
    args = parser.parse_args()

    if args.list:
        print("\n📋 可用测试用例:\n")
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name} — {tc.description}")
        print()
        return

    indices = list(range(1, len(TEST_CASES) + 1)) if args.test == "all" else [int(args.test)]
    ok = asyncio.run(
        run_tests(
            indices,
            auto_open=not args.no_auto_open,
            cleanup_on_success=not args.always_cleanup,
        )
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
