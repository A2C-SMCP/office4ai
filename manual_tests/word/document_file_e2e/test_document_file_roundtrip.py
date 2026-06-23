"""
Word whole-document .docx Round-Trip E2E Tests (自动化版本, #46-D)

测试 word_get_document_file / word_insert_document_file 工具对的整篇 round-trip。
这是 insertOoxml 吃不下整篇 Flat OPC 包后的整篇路：base64 .docx + insertFileFromBase64。

测试场景:
1. 整篇 round-trip 保真 — get_document_file → 落盘 .docx → insert_document_file(Replace, body)
   → 校验文档已知正文文本保真（双重验证：协议 + 文档内容）
2. 整篇 .docx 导出 — get_document_file 落盘，校验 .docx 是 zip 包（PK 魔数）、
   data 不回 inline base64、bytes > 0

设计要点:
- 驱动 MCP 工具（WordGetDocumentFileTool / WordInsertDocumentFileTool），file-handle 在工具内。
- 整篇 .docx 是二进制 zip，落盘/读取均为字节，服务端 base64 编解码（镜像 PPT）。
- F2 保真：insertFileFromBase64 是正文替换，文档级部件不保证 round-trip——只断言正文文本保真，
  不假设页眉页脚/节属性字节级对称。

⚠️ 外部依赖（真机）: 需 office-editor4ai Add-In 已实现 word:get/insert:documentFile handler（#63）。

运行方式:
    uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --test 1
    uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --test all
    uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --list
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
from office4ai.a2c_smcp.tools.word import WordGetDocumentFileTool, WordInsertDocumentFileTool
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "document_file_e2e"

# simple.docx 夹具中已知存在的文本（ensure_fixtures 生成），用于 round-trip 保真校验
KNOWN_TEXT_CANDIDATES = ["测试", "文本", "Hello", "段落"]

# .docx 是 zip 包，魔数 PK\x03\x04
ZIP_MAGIC = b"PK\x03\x04"


def _known_text_in(reader: DocumentReader) -> str | None:
    for candidate in KNOWN_TEXT_CANDIDATES:
        if reader.contains(candidate):
            return candidate
    return None


TEST_CASES: list[TestCase] = [
    TestCase(
        name="整篇 .docx round-trip 保真",
        fixture_name="simple.docx",
        description="get_document_file → 落盘 .docx → insert_document_file(Replace, body) → 校验正文文本保真",
        tags=["roundtrip"],
    ),
    TestCase(
        name="整篇 .docx 导出",
        fixture_name="simple.docx",
        description="get_document_file 落盘，校验 .docx 为 zip(PK 魔数)、data 不回 inline base64、bytes>0",
        tags=["export"],
    ),
]


async def _run_roundtrip(workspace: OfficeWorkspace, fixture: DocumentFixture) -> bool:
    """场景 1: 整篇 .docx round-trip 保真。"""
    get_tool = WordGetDocumentFileTool(workspace)
    insert_tool = WordInsertDocumentFileTool(workspace)
    dest = fixture.working_path.parent / "exported_document.docx"

    print("\n📝 步骤 1: word_get_document_file → 落盘 .docx...")
    get_result = await get_tool.execute({"document_uri": fixture.document_uri, "dest_path": str(dest)})
    if not get_result.get("success"):
        print(f"   ❌ get 失败: {get_result.get('error')}")
        return False
    if not dest.is_file() or dest.read_bytes()[:4] != ZIP_MAGIC:
        print(f"   ❌ dest_path 未落盘或非 .docx(zip): {dest}")
        return False
    print(f"   ✅ 导出成功，.docx {get_result['data'].get('bytes')} 字节 → {dest.name}")

    print("\n📝 步骤 2: word_insert_document_file(source_path, insertLocation=Replace, scope=body)...")
    insert_result = await insert_tool.execute(
        {
            "document_uri": fixture.document_uri,
            "source_path": str(dest),
            "insertLocation": "Replace",
            "scope": "body",
        }
    )
    if not insert_result.get("success"):
        print(f"   ❌ insert 失败: {insert_result.get('error')}")
        return False
    print("   ✅ insert 协议返回成功")

    # 双重验证：读回文档，已知正文文本应仍在（F2：只断言正文保真，不假设页眉页脚字节对称）
    print("\n📊 步骤 3: 正文内容保真校验(DocumentReader, F2 仅正文)...")
    reader = DocumentReader(fixture.working_path)
    reader.reload()
    hit = _known_text_in(reader)
    if hit is None:
        print(f"   ❌ round-trip 后未找到任何已知正文文本 {KNOWN_TEXT_CANDIDATES}")
        print(f"      文档前 200 字符: {reader.text[:200]}")
        return False
    print(f"   ✅ round-trip 后正文已知文本仍在: '{hit}'(页眉页脚/节属性按 F2 不做字节级断言)")
    return True


async def _run_export(workspace: OfficeWorkspace, fixture: DocumentFixture) -> bool:
    """场景 2: 整篇 .docx 导出（仅 get）。"""
    get_tool = WordGetDocumentFileTool(workspace)
    dest = fixture.working_path.parent / "exported_only.docx"

    print("\n📝 执行: word_get_document_file → 落盘...")
    result = await get_tool.execute({"document_uri": fixture.document_uri, "dest_path": str(dest)})
    if not result.get("success"):
        print(f"   ❌ get 失败: {result.get('error')}")
        return False

    data = result.get("data", {})
    passed = True
    if "base64" in data:
        print("   ❌ data 不应回传 inline base64")
        passed = False
    else:
        print("   ✅ data 未回传 inline base64(留在磁盘)")
    if not dest.is_file() or dest.read_bytes()[:4] != ZIP_MAGIC:
        print("   ❌ dest_path 未落盘或非 .docx(zip)")
        passed = False
    elif not data.get("bytes", 0) > 0:
        print(f"   ❌ bytes 非正: {data.get('bytes')}")
        passed = False
    else:
        print(f"   ✅ dest_path 落盘 .docx(zip)，bytes={data.get('bytes')}")
    return passed


_RUNNERS = [_run_roundtrip, _run_export]


async def run_single_test(runner: E2ETestRunner, test_case: TestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")
    print(f"📄 夹具: {test_case.fixture_name}")

    fixture_path = f"document_file_e2e/{test_case.fixture_name}"
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


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Word whole-document .docx Round-Trip E2E Tests (自动化版本, #46-D)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--test",
        choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"],
        default="1",
        help="要运行的测试: 1=整篇 round-trip, 2=整篇导出, all=全部",
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
    ok = asyncio.run(run_tests(indices, auto_open=not args.no_auto_open, cleanup_on_success=not args.always_cleanup))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
