# UAT 场景：Excel 工具注册验收 (excel-tools)

## 测试目标

验证 37 个 Excel MCP 工具在 MCP Inspector 中正确注册，参数 schema 符合预期。

> ⚠️ **前置：需 Excel Add-In 已连接**（W4a/#63 动态收敛）。Excel 工具 `requires_connection=True`，**断连态下不出现是正常收敛、非缺陷**（见 `tool-convergence` 场景 Round 1）。本场景为**连接态**下的逐工具 schema 核对。

> 约定：几乎所有工具都带可选参数 `worksheet_name`（省略则作用于活动工作表）。下表仅在该项为唯一/关键可选参数时标注，其余可选参数逐一列出。

## 验证清单

### 读取状态工具（3 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-01 | `excel_get_workbook_info` | 工作簿信息（工作表列表等） | `document_uri` | - |
| E-02 | `excel_get_worksheet_info` | 工作表信息（已用区域等） | `document_uri` | `worksheet_name` |
| E-03 | `excel_get_selected_range` | 当前选中区域 | `document_uri` | - |

### 区域工具（7 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-04 | `excel_get_range` | 读取区域值 | `document_uri`, `address` | `worksheet_name`, `include_format` |
| E-05 | `excel_set_range` | 写入区域值 | `document_uri`, `address`, `values` | `worksheet_name` |
| E-06 | `excel_clear_range` | 清除区域 | `document_uri`, `address`, `clear_type` | `worksheet_name` |
| E-07 | `excel_copy_range` | 复制区域 | `document_uri`, `source_address`, `target_address` | `worksheet_name` |
| E-08 | `excel_delete_range` | 删除区域（含移位） | `document_uri`, `address`, `shift_direction` | `worksheet_name` |
| E-09 | `excel_insert_range` | 插入区域（含移位） | `document_uri`, `address`, `shift_direction` | `worksheet_name` |
| E-10 | `excel_set_formula` | 设置公式 | `document_uri`, `address`, `formula` | `worksheet_name` |

### 格式工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-11 | `excel_get_range_format` | 读取区域格式 | `document_uri`, `address` | `worksheet_name` |
| E-12 | `excel_set_range_format` | 设置区域格式 | `document_uri`, `address`, `format` | `worksheet_name` |
| E-13 | `excel_add_conditional_format` | 添加条件格式 | `document_uri`, `address`, `rule` | `worksheet_name` |
| E-14 | `excel_clear_conditional_format` | 清除条件格式 | `document_uri`, `address` | `worksheet_name` |
| E-15 | `excel_merge_cells` | 合并单元格 | `document_uri`, `address` | `worksheet_name` |
| E-16 | `excel_unmerge_cells` | 取消合并 | `document_uri`, `address` | `worksheet_name` |

### 工作表工具（5 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-17 | `excel_get_worksheets` | 工作表列表 | `document_uri` | - |
| E-18 | `excel_add_worksheet` | 新增工作表 | `document_uri` | `name` |
| E-19 | `excel_delete_worksheet` | 删除工作表 | `document_uri`, `worksheet_name` | - |
| E-20 | `excel_rename_worksheet` | 重命名工作表 | `document_uri`, `current_name`, `new_name` | - |
| E-21 | `excel_activate_worksheet` | 激活工作表 | `document_uri`, `worksheet_name` | - |

### 表格工具（6 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-22 | `excel_insert_table` | 插入表格 | `document_uri`, `address`, `has_headers` | `data`, `style_name`, `worksheet_name` |
| E-23 | `excel_get_table` | 获取单个表格 | `document_uri`, `table_id` | `worksheet_name` |
| E-24 | `excel_get_tables` | 获取表格列表 | `document_uri` | `worksheet_name` |
| E-25 | `excel_add_table_row` | 追加表格行 | `document_uri`, `table_id`, `values` | `worksheet_name` |
| E-26 | `excel_delete_table_row` | 删除表格行 | `document_uri`, `table_id`, `row_index` | `worksheet_name` |
| E-27 | `excel_sort_table` | 排序表格 | `document_uri`, `table_id`, `sort_fields` | `worksheet_name` |

### 图表工具（4 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-28 | `excel_insert_chart` | 插入图表 | `document_uri`, `source_address`, `chart_type` | `title`, `position`, `worksheet_name` |
| E-29 | `excel_get_charts` | 获取图表列表 | `document_uri` | `worksheet_name` |
| E-30 | `excel_update_chart` | 更新图表 | `document_uri`, `chart_name`, `properties` | `worksheet_name` |
| E-31 | `excel_delete_chart` | 删除图表 | `document_uri`, `chart_name` | `worksheet_name` |

### 数据透视表工具（3 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-32 | `excel_insert_pivot_table` | 插入数据透视表 | `document_uri`, `source_address`, `target_address` | `name`, `worksheet_name` |
| E-33 | `excel_get_pivot_tables` | 获取数据透视表列表 | `document_uri` | `worksheet_name` |
| E-34 | `excel_delete_pivot_table` | 删除数据透视表 | `document_uri`, `pivot_table_name` | `worksheet_name` |

### 查找与筛选工具（3 个）

| # | 工具名 | 描述关键词 | 必填参数 | 可选参数 |
|---|--------|-----------|----------|----------|
| E-35 | `excel_find_values` | 查找值 | `document_uri`, `search_text` | `address`, `worksheet_name`, `match_case`, `match_entire_cell` |
| E-36 | `excel_set_auto_filter` | 设置自动筛选 | `document_uri`, `address`, `criteria` | `worksheet_name` |
| E-37 | `excel_clear_auto_filter` | 清除自动筛选 | `document_uri` | `worksheet_name` |

## 逐项验证要点

每个工具检查：

1. **存在性**：工具出现在 MCP Inspector Tools 列表中
2. **名称**：`excel_` 前缀，snake_case 命名
3. **描述**：非空，能让用户理解工具用途
4. **`document_uri`**：必填，类型 string
5. **其他参数**：必填/可选与上表一致，类型合理（`worksheet_name` 恒为可选，省略作用于活动表）
6. **总数**：Excel 工具共 **37** 个，与本清单一一对应，无遗漏、无多余
