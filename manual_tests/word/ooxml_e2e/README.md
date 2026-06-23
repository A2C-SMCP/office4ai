# Word OOXML Round-Trip E2E Tests

`word:get:ooxml` / `word:insert:ooxml` 事件对的端到端测试（自动化版本，OASP /word Draft v0.3.0，office4ai#46）。

驱动 MCP 工具 `word_get_ooxml` / `word_insert_ooxml`（file-handle 读写逻辑在工具内），做 OOXML **片段** round-trip。

> ⚠️ **路径收窄（#46-D）**：`word_insert_ooxml` 仅吃**片段**；整篇文档包会被 Office.js `insertOoxml` 以 `GeneralException` 拒绝。**整篇 round-trip 已移到 [`../document_file_e2e/`](../document_file_e2e/)**（base64 `.docx` + `insertFileFromBase64`）。本套只覆盖片段/选区语义。

## ⚠️ 外部依赖（真机阻塞）

真机运行需 **office-editor4ai (word-editor4ai) Add-In** 已实现 `word:get:ooxml` / `word:insert:ooxml` 两个 handler。
双方已通过 cross-ask **锁定 wire 契约**（见 office4ai#46）：

```
word:get:ooxml    req { scope?: "selection"|"body" = "selection" }
                  res data { scope(生效值，空选区回退 body 时为 "body"), ooxml(Flat OPC XML 字符串) }
word:insert:ooxml req { ooxml(Flat OPC XML 字符串，非 base64), insertLocation: Replace|Start|End, scope?: selection|body }
                  res data { scope, insertLocation }
```

Add-In 就绪前，`get` 阶段会超时或返回 `3016 API_NOT_SUPPORTED`。脚手架本身（结构 / import / `--list`）不依赖 Add-In。

## 测试文件

### test_ooxml_roundtrip.py

| 测试编号 | 测试名称 | 描述 |
|---------|---------|------|
| 1 | 选区片段 round-trip | `get(selection)` → 落盘片段 Flat OPC → `insert(Replace, selection)` → 校验正文保真（需先在 Word 手动选中一段文字；空选区会提示改用 `document_file_e2e`） |
| 2 | Body OOXML 只读导出 | `get(body)` 落盘，校验 Flat OPC、`data` 不回 inline `ooxml`、`scope` 回生效值（整篇包仅可读，**不**经 `word_insert_ooxml` 回灌） |

## 运行方式

### 前置条件

1. 在 Word 中加载 office-editor4ai Add-In（且 Add-In 已实现 OOXML handler）
2. 测试会自动启动 Workspace 服务器、打开文档并等待 Add-In 连接

### 命令

```bash
# 列出用例（不需 Add-In）
uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --list

# 单个测试
uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --test 1
uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --test 2

# 全部
uv run python manual_tests/word/ooxml_e2e/test_ooxml_roundtrip.py --test all
```

夹具由 `ensure_fixtures()` 自动生成于 `manual_tests/word/fixtures/ooxml_e2e/`（`simple.docx` 等标准夹具）。

## 验证设计（双重验证）

- **协议验证**：`get` 返回 `{scope, filePath, bytes}` 且不回 inline `ooxml`；`insert` 返回 `{scope, insertLocation}`。
- **内容验证**：round-trip 后用 `DocumentReader` 读回 `.docx`，断言已知文本仍在（保真未丢内容）。
