# UAT 场景：Word 功能验收 (word-functional) — Phase 2

## 测试目标

在真实 Word + Add-In 环境下，验证 Word 工具**功能可正常跑通**——由 `manual_tests/word/` 的 E2E 用例驱动，python-docx 双重验证 OOXML。触发人协助连接 Add-In，逐个走查。

> Phase 1（`word-tools` 注册验收，25 工具）确认「工具在」；本场景确认「工具能用」。

## 前置

- macOS + Word 桌面版；office-editor4ai `word-editor4ai` dev server 运行（`pnpm dev:word`）。
- 触发人：打开 Word 文档后**手动激活 taskpane**（加载项面板点 word-editor）握手成功。
- 部分用例需**按脚本提示手动操作**（如 `get_selection` / `replace_selection` / `select_text` 会提示「请在 Word 中选中某段文本」）——触发人照做。
- 运行需 `dangerouslyDisableSandbox`；category runner 连接超时 30s，需及时激活 taskpane。

## Health 档（接线桥接，**无需 Add-In**）

```bash
uv run python manual_tests/word/test_word_e2e.py --mode health         # 基础事件接线
uv run python manual_tests/word/test_word_table_e2e.py --mode health   # 4 表格事件接线
```

**通过标准**：事件均已注册、无缺失。

## Smoke 档（默认，关键类别）

| # | 用例 | 命令（`uv run python manual_tests/word/…`） | 人工交互 |
|---|---|---|---|
| WS-1 | 插入文本 | `insert_text_e2e/test_basic_insert.py --test all` | 激活 taskpane |
| WS-2 | 替换文本 | `replace_text_e2e/test_basic_replace.py --test all` | 激活 taskpane |
| WS-3 | 选中文本 | `select_text_e2e/test_basic_select.py --test all` | 激活 + 按提示选文本 |
| WS-4 | 评论 CRUD | `comment_e2e/test_comment_crud.py --test all` | 激活 taskpane |
| WS-5 | 表格 4 工具 | `test_word_table_e2e.py --mode tables` | 激活 taskpane |

**通过标准**：5 项全 ✅（协议成功 + python-docx 内容双验）。

## Full 档（逐类别深走，11 类 + 表格）

| 类别 | 代表命令（`--test all`；多文件逐个跑） | 覆盖 |
|---|---|---|
| get_visible_content | `get_visible_content_e2e/test_basic_get.py` · `test_options_get.py` · `test_edge_cases.py` | 可见内容读取 |
| get_selection | `get_selection_e2e/test_selection.py` | 选区信息（需选文本） |
| get_document_structure | `get_document_structure_e2e/test_basic_structure.py` | 大纲/结构 |
| get_document_stats | `get_document_stats_e2e/test_basic_stats.py` | 文档统计 |
| get_styles | `get_styles_e2e/test_styles.py` | 样式列表 |
| insert_text | `insert_text_e2e/test_basic_insert.py` · `test_format_insert.py` · `test_location_insert.py` | 插入（基础/格式/定位）|
| replace_text | `replace_text_e2e/test_basic_replace.py` · `test_format_replace.py` | 替换 |
| replace_selection | `replace_selection_e2e/test_text_replace.py` · `test_format_replace.py` · `test_edge_cases.py` | 替换选区（需选文本）|
| select_text | `select_text_e2e/test_basic_select.py` · `test_search_options.py` · `test_selection_modes.py` · `test_edge_cases.py` | 选中（选项/模式）|
| export_content | `export_content_e2e/test_basic_export.py` · `test_export_options.py` | 导出 |
| comment | `comment_e2e/test_comment_crud.py` · `test_comment_options.py` · `test_comment_target.py` | 评论 CRUD/选项/定位 |
| 表格 | `test_word_table_e2e.py --mode tables`（`--mode b1` 补边界） | merge/update cell·row·col·format |

（命令前缀均为 `uv run python manual_tests/word/…  --test all`）

**通过标准**：每类别用例全 ✅；失败项记录实际现象与提示。

## 报告片段

```
### Word 功能验收
- Health（事件接线）: ✅
- Smoke（WS-1..5）: 5/5 ✅
- Full（11 类 + 表格）: get_* ✅ / insert_text ✅ / replace_* ✅ / select_text ✅ / export ✅ / comment ✅ / 表格 ✅
```
