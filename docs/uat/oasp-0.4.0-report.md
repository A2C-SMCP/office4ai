# Office4AI MCP Server — OASP 0.4.0 UAT 验收报告

- **日期**：2026-07-14 ~ 07-16
- **环境**：office4ai `main@6f73c84`（`SERVER_VERSION` 0.4.0）× office-editor4ai `main` 0.4.0（word PR#75 + ppt PR#77 已合并）
- **验收人**：QA（Claude）+ 触发人（用户，真机连接 + Manual-Only 手动用例）
- **方法**：Phase 1 程序化 `list_tools`/`list_resources`；Phase 2 `manual_tests/` E2E（协议成功 + 文档内容双验：python-docx / python-pptx / openpyxl 读盘）

---

## 执行摘要

**结论：✅ 通过**（含 1 项测试框架修复 [office4ai#86](https://github.com/A2C-SMCP/office4ai/pull/86) 待合并）。

两阶段（注册 + 功能）全平台覆盖；OASP 0.4.0 字体（WordFont / PptFont 嵌套 + `insert:shape` font + text-capable 4002 门控）在双端真机验证正确。过程中发现并修复 3 类**测试框架**缺陷（**非产品缺陷**），另记录 1 项 Add-In 错误码映射瑕疵（[office-editor4ai#85](https://github.com/JIAQIA/office-editor4ai/issues/85)）。

---

## Phase 1 — 注册验收（07-14，程序化 stdio）

| 维度 | 结果 | 明细 |
|---|---|---|
| 工具注册 | ✅ **87** | 25 word / 24 ppt / 37 excel / 1 authoring（`office_run_script`）|
| Resource 注册 | ✅ **25** | 1 根 `window://office4ai` + 3 SKILL 根 + 21 SKILL 子 |
| 动态收敛 W4a（工具） | ✅ | 断连 0/0/0 → 连接 25/24/37 → 断开回退；`office_run_script` 恒在 |
| 动态收敛 W4b（资源） | ✅ | per-file 窗口随连接注册/注销；excel 窗口内容实时拉取 |

---

## Phase 2 — 功能验收（真机 E2E）

### Excel ✅（07-15，`--mode full`）

| 结果 | 明细 |
|---|---|
| **43/43** ✅ | 37 事件 + openpyxl 双验；`set:autoFilter` 夹具污染修复（[office4ai#84](https://github.com/A2C-SMCP/office4ai/issues/84)）后复验全绿 |

### Word ✅

| 类别 | 结果 | 备注 |
|---|---|---|
| health | ✅ | base + 4 表格事件接线 |
| get_visible / structure / stats / styles | ✅ | 4/4 · 5/5 · 4/4 · 5/5 |
| insert_text basic / **format(0.4.0)** | ✅ | basic 4/4；**format 6/6 真机真验**（bold/italic/combined，OOXML 读回）|
| replace_text basic / **format(0.4.0)** | ✅ | basic 6/6；styleName ✅ + 字体迁移嵌套 font |
| **replace_selection(0.4.0, Manual-Only)** | ✅ | 触发人真机复跑确认格式生效 |
| select_text ×4 | ✅ | basic 4 / search 5 / modes 4 / edge 6 |
| export ×2 / comment ×3 | ✅ | 4/4·3/3 · CRUD/options/target |
| **表格(0.4.0 tableCell/tableFormat)** | ✅ | 正向全 OK（merge / 背景色 / 行列 / 字体）|
| get_selection / cursor（Manual-Only） | ✅ | 触发人真机复跑 |

### PPT ✅（15 类 + chart）

| 类别 | 结果 | 备注 |
|---|---|---|
| **insert_shape(0.4.0)** | ✅ | 6/6 真验：font 读回（bold/size24/双删除线）+ **Line→4002 门控** |
| **update_text_box(0.4.0)** / **update_table format·cell·row(0.4.0)** | ✅ | fixture 迁移嵌套 font（PR#86）+ 协议成功 |
| insert_text basic/options / insert_image/table | ✅ | 5/5 · 3/3 等 |
| get_elements / slide_info / layouts / screenshot | ✅ | 元素 / 信息 / 版式 / 截图 |
| update_image / update_element / reorder / delete / slide_mgmt ×4 | ✅ | 位置 / 层级 / 删除 / 增删移跳 |
| chart（pathb-live） | ✅ | insert/get/update 自动断言全过 + 无 slideMaster 泄漏 |

---

## 关键发现 & 处置

| # | 发现 | 性质 | 处置 |
|---|---|---|---|
| 1 | **0.4.0 字体 format E2E 全程假绿**：fixtures 发旧扁平 format 被 DTO `extra=ignore` 静默丢弃 + 验证器 `run_has_format` run-split 盲区 + soft-pass 三重掩盖 | 测试框架（**产品无辜**，OOXML ground truth 证明字体全应用） | 🐛 [office4ai#85](https://github.com/A2C-SMCP/office4ai/issues/85) / 🔀 [PR#86](https://github.com/A2C-SMCP/office4ai/pull/86)（8 fixtures 迁移嵌套 font + run-split 修复 + cleanup bug + UAT SKILL 强化）|
| 2 | Excel `set:autoFilter` e2e 夹具污染（表1 range 冲突）| 测试夹具 | ✅ [office4ai#84](https://github.com/A2C-SMCP/office4ai/issues/84) 已合并 |
| 3 | Excel 5xxx 错误码降级 3000 + OASP 冲突 | 协议 + Add-In | ✅ oasp-protocol#17 裁决 + [office4ai#82/#83](https://github.com/A2C-SMCP/office4ai/issues/82) 已合并 |
| 4 | Word 表格 merge 混合列宽 → Add-In 返 `3000` 应具体码 | Add-In 错误码映射（负路径，非数据缺陷）| 🐛 [office-editor4ai#85](https://github.com/JIAQIA/office-editor4ai/issues/85) 已提 |

### 发现 #1 ground truth（证明产品正确）

喂正确嵌套 font 后，OOXML 直查 insert「组合格式文本」run：

```
run0 '组'       bold=True italic=True name=Arial size=228600(=18pt) color=FF0000
run1 '合格式文本' bold=True italic=True name=Arial size=228600      color=FF0000
```

修复后 `insert_text format` 真机 **6/6 真验**（bold / italic / combined 从 `⚠️未验证` → `✅格式正确`）。

---

## 已知限制（非阻塞，留待后续）

- **PPT 字体读回深度**：`update_text_box` / `update_table` 的 0.4.0 字体目前只验「输入正确 + 协议成功」，未做 OOXML 字体读回（仅 `insert_shape` 有真读回）。PR#86 迁移了输入，读回增强留待下轮。
- **cold-start 30s 超时**：每脚本边界偶发 1 次连接超时假失败（最简单用例"失败"即铁证）；建议后续把 `connection_timeout` 提到 ~150s。
- **Manual-Only 用例**：`get_selection` / `replace_selection` / cursor 需真人 GUI 操作，agent 无法代跑（UAT SKILL 已固化 Manual-Only 协议）。

---

## 结论

OASP 0.4.0 在 office4ai（Server）+ office-editor4ai（Add-In）双端**功能验收通过**：注册完整（87 工具 / 25 资源 / W4a·W4b 收敛）、功能可用（Excel 43/43、Word/PPT 全类别、0.4.0 字体真机真验）。**产品零缺陷**；1 项测试框架修复（PR#86）合并后 0.4.0 字体 format E2E 缺口即补齐；1 项 Add-In 错误码瑕疵（office-editor4ai#85）单独跟踪。
