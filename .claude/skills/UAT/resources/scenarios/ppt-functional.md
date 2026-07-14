# UAT 场景：PPT 功能验收 (ppt-functional) — Phase 2

## 测试目标

在真实 PowerPoint + Add-In 环境下，验证 PPT 工具**功能可正常跑通**——由 `manual_tests/ppt/` 的 E2E 用例驱动，python-pptx 双重验证 OOXML。触发人协助连接 Add-In，逐个走查。

> Phase 1（`ppt-tools` 注册验收，24 工具）确认「工具在」；本场景确认「工具能用」。

## 前置

- macOS + PowerPoint 桌面版；office-editor4ai `ppt-editor4ai` dev server 运行（`pnpm dev:ppt`）。
- 触发人：打开 pptx 后**手动激活 taskpane**（加载项面板点 ppt-editor）握手成功。
- 运行需 `dangerouslyDisableSandbox`；category runner 连接超时 30s，需及时激活 taskpane。
- chart 类由 Server 端 python-pptx 直接改盘（离线路径已删，走 Add-In），见其脚本 `--mode`。

## Health 档（接线桥接）

PPT 无专用 `--mode health` 驱动；事件接线回归由 unit + contract 测试守护（`poe test`）。chart 类真机走 `--mode pathb-live`（见下）。

## Smoke 档（默认，关键类别）

| # | 用例 | 命令（`uv run python manual_tests/ppt/…`） | 人工交互 |
|---|---|---|---|
| PS-1 | 插入文本框 | `insert_text_e2e/test_basic_insert.py --test all` | 激活 taskpane |
| PS-2 | 插入形状（含 font, 0.4.0）| `insert_shape_e2e/test_shape_insert.py --test all` | 激活 taskpane |
| PS-3 | 幻灯片管理 | `slide_management_e2e/test_add_slide.py --test all`（+ delete/move/goto）| 激活 taskpane |
| PS-4 | 更新文本框 | `update_text_box_e2e/test_text_box_update.py --test all` | 激活 taskpane |
| PS-5 | 图表 | `test_chart_e2e.py --mode pathb-live`（真机真实 Add-In）| 激活 taskpane |

> PS-2 含本轮新增的 insert-with-font 读回校验 + Line→4002 负例（office4ai#76），是重点回归项。

**通过标准**：5 项全 ✅（协议成功 + python-pptx 内容双验）。

## Full 档（逐类别深走，15 类 + 图表）

| 类别 | 代表命令（`--test all`；多文件逐个跑） | 覆盖 |
|---|---|---|
| get_elements | `get_elements_e2e/test_current_slide.py` · `test_specific_slide.py` | 元素枚举 |
| get_slide_info | `get_slide_info_e2e/test_basic_info.py` | 幻灯片信息 |
| get_slide_layouts | `get_slide_layouts_e2e/test_layouts.py` | 版式 |
| get_screenshot | `get_screenshot_e2e/test_screenshot.py` | 截图 |
| insert_text | `insert_text_e2e/test_basic_insert.py` · `test_insert_options.py` | 插入文本框（含 font）|
| insert_image | `insert_image_e2e/test_image_insert.py` | 插入图片 |
| insert_table | `insert_table_e2e/test_table_insert.py` | 插入表格 |
| insert_shape | `insert_shape_e2e/test_shape_insert.py` · `test_shape_types.py` | 插入形状（含 font/Line→4002）|
| update_text_box | `update_text_box_e2e/test_text_box_update.py` | 更新文本框 |
| update_image | `update_image_e2e/test_image_update.py` | 更新图片 |
| update_table | `update_table_e2e/test_cell_update.py` · `test_row_column_update.py` · `test_table_format.py` | 表格单元格/行列/格式 |
| update_element | `update_element_e2e/test_element_update.py` | 位置/大小/旋转 |
| reorder_element | `reorder_element_e2e/test_reorder.py` | 层级重排 |
| delete_element | `delete_element_e2e/test_delete.py` | 删除元素 |
| slide_management | `slide_management_e2e/test_add_slide.py` · `test_delete_slide.py` · `test_move_slide.py` · `test_goto_slide.py` | 幻灯片增删移跳 |
| 图表 | `test_chart_e2e.py`（按 `--mode`）| insert/get/update chart |

（命令前缀均为 `uv run python manual_tests/ppt/…  --test all`）

**通过标准**：每类别用例全 ✅；失败项记录实际现象。

## 报告片段

```
### PPT 功能验收
- Smoke（PS-1..5）: 5/5 ✅
- Full（15 类 + 图表）: get_* ✅ / insert_*(text/image/table/shape) ✅ / update_*(textbox/image/table/element) ✅ / reorder ✅ / delete ✅ / slide_mgmt ✅ / chart ✅
```
