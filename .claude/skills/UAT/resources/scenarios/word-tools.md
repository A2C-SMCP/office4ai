# UAT 场景：Word 工具注册验收 (word-tools)

## 测试目标

验证 25 个 Word MCP 工具在 MCP Inspector 中正确注册，参数 schema 符合预期。

> ⚠️ **前置：需 Word Add-In 已连接**（W4a/#63 动态收敛）。Word 工具 `requires_connection=True`，**断连态下不出现是正常收敛、非缺陷**（见 `tool-convergence` 场景 Round 1）。本场景为**连接态**下的逐工具 schema 核对。

## 验证清单

### 读取工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-01 | `word_get_selected_content` | 选中内容 | `document_uri` | `options` |
| W-02 | `word_get_visible_content` | 可见内容 | `document_uri` | `options` |
| W-03 | `word_get_selection` | 选区信息 | `document_uri` | - |
| W-04 | `word_get_document_structure` | 文档结构/大纲 | `document_uri` | - |
| W-05 | `word_get_document_stats` | 文档统计 | `document_uri` | - |
| W-06 | `word_get_styles` | 样式列表 | `document_uri` | `options` |

### 文本操作工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-07 | `word_insert_text` | 插入文本 | `document_uri`, `text` | `location`, `format` |
| W-08 | `word_append_text` | 追加文本 | `document_uri`, `text` | `location`, `format` |
| W-09 | `word_replace_text` | 替换文本 | `document_uri`, `search_text`, `replace_text` | `options`, `format` |
| W-10 | `word_replace_selection` | 替换选区 | `document_uri`, `content` | - |
| W-11 | `word_select_text` | 选中文本 | `document_uri`, `search_text` | `search_options`, `selection_mode`, `select_index` |

### 多媒体工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-12 | `word_insert_image` | 插入图片 | `document_uri`, `image` | `location`, `wrap_type` |
| W-13 | `word_insert_table` | 插入表格 | `document_uri`, `options` | - |
| W-14 | `word_insert_equation` | 插入公式（LaTeX） | `document_uri`, `latex` | `options` |
| W-15 | `word_insert_toc` | 插入目录 | `document_uri` | `options` |

### 表格操作工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-16 | `word_merge_cells` | 合并单元格 | `document_uri`, `start_row_index`, `start_column_index`, `end_row_index`, `end_column_index` | `table_id` |
| W-17 | `word_update_table_cell` | 更新单元格内容 | `document_uri`, `cells` | `table_id` |
| W-18 | `word_update_table_row_column` | 增删行/列 | `document_uri` | `table_id`, `rows`, `columns` |
| W-19 | `word_update_table_format` | 表格格式 | `document_uri` | `table_id`, `style_options`, `border_options`, `column_widths`, `alignment` |

### 导出工具（1 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-20 | `word_export_content` | 导出内容 | `document_uri`, `format` | `options` |

### 评论工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| W-21 | `word_get_comments` | 获取评论 | `document_uri` | `options` |
| W-22 | `word_insert_comment` | 插入评论 | `document_uri`, `text` | `target` |
| W-23 | `word_delete_comment` | 删除评论 | `document_uri`, `comment_id` | - |
| W-24 | `word_reply_comment` | 回复评论 | `document_uri`, `comment_id`, `text` | - |
| W-25 | `word_resolve_comment` | 解决评论 | `document_uri`, `comment_id` | `resolved` |

## 逐项验证要点

每个工具检查：

1. **存在性**：工具出现在 MCP Inspector Tools 列表中
2. **名称**：`word_` 前缀，snake_case 命名
3. **描述**：非空，能让用户理解工具用途
4. **`document_uri`**：必填，类型 string
5. **其他参数**：必填/可选与上表一致，类型合理
6. **总数**：Word 工具共 **25** 个，与本清单一一对应，无遗漏、无多余
