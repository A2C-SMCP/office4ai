# UAT 场景：PPT 工具注册验收 (ppt-tools)

## 测试目标

验证 24 个 PPT MCP 工具在 MCP Inspector 中正确注册，参数 schema 符合预期。

> ⚠️ **前置：需 PPT Add-In 已连接**（W4a/#63 动态收敛）。PPT 工具 `requires_connection=True`，**断连态下不出现是正常收敛、非缺陷**（见 `tool-convergence` 场景 Round 1）。本场景为**连接态**下的逐工具 schema 核对。
>
> 注：PPT 工具入参用 **camelCase**（如 `slideIndex` / `elementId` / `shapeType`），与 word/excel 的 snake_case 不同——这是实际 schema，照此核对。

## 验证清单

### 内容读取工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-01 | `ppt_get_current_slide_elements` | 当前幻灯片元素 | `document_uri` | - |
| P-02 | `ppt_get_slide_elements` | 指定幻灯片元素 | `document_uri`, `slideIndex` | `options` |
| P-03 | `ppt_get_slide_screenshot` | 幻灯片截图 | `document_uri`, `slideIndex` | `options` |
| P-04 | `ppt_get_slide_info` | 幻灯片信息 | `document_uri` | `slideIndex` |
| P-05 | `ppt_get_slide_layouts` | 版式列表 | `document_uri` | `options` |

### 内容插入工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-06 | `ppt_insert_text` | 插入文本框 | `document_uri`, `text` | `options` |
| P-07 | `ppt_insert_image` | 插入图片 | `document_uri`, `image` | `options` |
| P-08 | `ppt_insert_table` | 插入表格 | `document_uri`, `options` | - |
| P-09 | `ppt_insert_shape` | 插入形状（含 font, 0.4.0） | `document_uri`, `shapeType` | `options` |

### 元素更新工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-10 | `ppt_update_text_box` | 更新文本框（含 font） | `document_uri`, `elementId`, `updates` | - |
| P-11 | `ppt_update_image` | 更新图片 | `document_uri`, `elementId`, `image` | `options` |
| P-12 | `ppt_update_table_cell` | 更新表格单元格 | `document_uri`, `elementId`, `cells` | - |
| P-13 | `ppt_update_table_row_column` | 增删表格行/列 | `document_uri`, `elementId` | `rows`, `columns` |
| P-14 | `ppt_update_table_format` | 表格格式（含 font/对齐） | `document_uri`, `elementId` | `cellFormats`, `rowFormats`, `columnFormats` |
| P-15 | `ppt_update_element` | 更新元素（位置/大小/旋转等） | `document_uri`, `elementId`, `updates` | `slideIndex` |

### 图表工具（3 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-16 | `ppt_insert_chart` | 插入图表 | `document_uri`, `chart` | `options` |
| P-17 | `ppt_get_chart` | 获取图表 | `document_uri`, `elementId` | `slideIndex` |
| P-18 | `ppt_update_chart` | 更新图表 | `document_uri`, `elementId`, `chart` | `slideIndex` |

### 元素管理工具（2 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-19 | `ppt_delete_element` | 删除元素 | `document_uri` | `elementId`, `elementIds`, `slideIndex` |
| P-20 | `ppt_reorder_element` | 层级重排（置顶/底等） | `document_uri`, `elementId`, `action` | `slideIndex` |

### 幻灯片管理工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| P-21 | `ppt_add_slide` | 新增幻灯片 | `document_uri` | `options` |
| P-22 | `ppt_delete_slide` | 删除幻灯片 | `document_uri`, `slideIndex` | - |
| P-23 | `ppt_move_slide` | 移动幻灯片 | `document_uri`, `fromIndex`, `toIndex` | - |
| P-24 | `ppt_goto_slide` | 跳转幻灯片 | `document_uri`, `slideIndex` | - |

## 逐项验证要点

每个工具检查：

1. **存在性**：工具出现在 MCP Inspector Tools 列表中
2. **名称**：`ppt_` 前缀，snake_case 命名
3. **描述**：非空，能让用户理解工具用途
4. **`document_uri`**：必填，类型 string
5. **其他参数**：必填/可选与上表一致（**入参键为 camelCase**），类型合理
6. **总数**：PPT 工具共 **24** 个，与本清单一一对应，无遗漏、无多余
