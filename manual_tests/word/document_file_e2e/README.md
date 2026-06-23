# Word whole-document .docx Round-Trip E2E Tests (#46-D)

`word:get:documentFile` / `word:insert:documentFile` 事件对的端到端测试（OASP /word Draft v0.3.0，office4ai#46-D）。

驱动 MCP 工具 `word_get_document_file` / `word_insert_document_file`，做**整篇** `.docx` round-trip。

## 为什么有这套（路径背景）

`getOoxml(body)` 产出的整篇 Flat OPC 包**无法**经 `insertOoxml` 回灌——Office.js 抛 `GeneralException`。
故整篇 round-trip 改走 **base64 `.docx` + `insertFileFromBase64`**（与 PPT 同路子）。片段/选区仍走 [`../ooxml_e2e/`](../ooxml_e2e/)。

## ⚠️ 外部依赖（真机阻塞）

需 **office-editor4ai (word-editor4ai) Add-In** 已实现 `word:get:documentFile` / `word:insert:documentFile` 两个 handler（cross-ask 重锁契约，#63）：

```
word:get:documentFile     req { requestId, documentUri } → res data { base64 }            # 整篇 .docx
word:insert:documentFile  req { requestId, documentUri, base64, insertLocation: Replace|Start|End, scope: body|selection }
                          → res data { scope, insertLocation }                            # 最小返回
```

- Add-In 端 `getFileAsync(Compressed)` 装配 + `Body/Range.insertFileFromBase64`。
- **F2 保真**：`insertFileFromBase64` 是**正文替换**，文档级部件（页眉页脚 / 文档级 sectPr / docProps / customXml）**不保证**完整 round-trip——本套只断言**正文文本保真**，不假设字节级对称。
- **F1 传输**：整篇 .docx base64 常 >1MB；office4ai 暂不预防性改 `max_http_buffer_size`，路 B 首次真机即作传输上限的经验测试（撞墙再定位真实限制）。

脚手架本身（结构 / import / `--list`）不依赖 Add-In。

## 测试文件

### test_document_file_roundtrip.py

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 整篇 .docx round-trip 保真 | `get_document_file` → 落盘 `.docx` → `insert_document_file(Replace, body)` → 校验正文文本保真（双重验证：协议 + 文档内容，F2 仅正文） |
| 2 | 整篇 .docx 导出 | `get_document_file` 落盘，校验 `.docx` 为 zip（PK 魔数）、`data` 不回 inline `base64`、`bytes` > 0 |

## 运行方式

```bash
uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --list
uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --test 1
uv run python manual_tests/word/document_file_e2e/test_document_file_roundtrip.py --test all
```

夹具由 `ensure_fixtures()` 自动生成于 `manual_tests/word/fixtures/document_file_e2e/`。
