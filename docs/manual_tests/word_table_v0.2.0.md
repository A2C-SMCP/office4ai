# Word 表格 4 个 MCP Tools 手工验收清单 (Issue #8 / OASP v0.2.0)

> **范围**：`word_merge_cells` / `word_update_table_cell` / `word_update_table_row_column` / `word_update_table_format`
> **协议引用**：https://doc.turingfocus.cn/oasp/0.2.0/specification/events-word/
> **关联**：JIAQIA/office4ai#8 · JIAQIA/office-editor4ai#23 (PR #27) · JIRA OF4AI-10 / OF4AI-11

本清单镜像 Issue #8 「手工测试清单（端到端视觉验证）」11 项，划分为：

- **A. 自动化部分**：跑 `test_word_table_e2e.py --mode tables`，脚本完成 OOXML 结构验证；你目测视觉
- **B. 手工部分**：3 个错误码场景 + AI 业务闭环，必须手工触发
- **C. 跨仓注意**：Server 侧未做的事

---

## 前置准备

```bash
# 1. 确保 office4ai 测试套件已通过
uv run poe pre-commit            # 应该 ✅ 887 passed

# 2. 确保已生成证书并安装到系统信任库
uv run office4ai-mcp setup       # 仅首次

# 3. 启动 office-editor4ai Add-In（另一个终端）
cd /Users/jqq/WebstormProjects/office-editor4ai/word-editor4ai
pnpm start                       # 或 npm start
```

确认：
- [ ] office4ai 单元 + 契约测试全部通过（`poe pre-commit`）
- [ ] office-editor4ai 已合入 PR #27（含 4 个新事件 Handler + `documentStructure.tables[]` 字段）
- [ ] Add-In dev server 在跑

---

## A. 自动化场景（跑脚本 + 你目测）

### A.0 端到端连通性 (Issue #8 项 1)

```bash
uv run python manual_tests/word/test_word_table_e2e.py --mode health
```

**预期**：
```
✅ Workspace 运行正常
✅ 4 个新表格事件均已注册到 request_registry: [
  'word:merge:cells', 'word:update:tableCell',
  'word:update:tableRowColumn', 'word:update:tableFormat'
]
```

- [ ] **A.0** Workspace 启动成功，4 个事件注册到 `request_registry`

---

### A.1～A.5 表格流水线（Issue #8 项 2/3/4/5/6 合并）

```bash
uv run python manual_tests/word/test_word_table_e2e.py --mode tables
```

脚本将：

1. 启动 Workspace + 自动激活 Word Add-In
2. 复制 `manual_tests/fixtures/empty.docx` 为工作副本并打开
3. 依次调用（顺序经过 Word.js API 限制重排）：
   1. `insert_table 5×4`
   2. **`update_table_format(columnWidths=[120,80,80,80])` — 必须在 merge 之前**
      （Word.js `TableColumnCollection` 在合并表上不可访问，详见
      [office-editor4ai#29](https://github.com/JIAQIA/office-editor4ai/issues/29) /
      [PR #31](https://github.com/JIAQIA/office-editor4ai/pull/31)）
   3. `merge_cells`（首行 4 列合并）
   4. `update_table_cell`（蓝底白字加粗设置首行表头）
   5. `update_table_row_column`（4 行数据）
   6. `update_table_cell`（标签列灰底加粗）
   7. `update_table_format(borders + cellPadding + alignment="Centered")` — 这些在合并表上仍可用
4. 调用一次缺省 `tableId` 的 `update_table_cell` 验证「光标解析」路径
5. 触发 Word 保存，用 python-docx 读取 OOXML 结构验证：
   - 表格存在
   - 首行被合并（所有底层 `_tc` 指向同一节点）
   - 首行单元格背景色 = `#1F4E79`
   - 第 2 行第 0/1 列 = "甲方" / "ACME Corp"
   - `tblGrid` 列数 = 4

> ⚠️ **Word.js API 限制（Microsoft 平台级硬限制）**：含合并单元格的表上**无法**通过
> `TableColumnCollection` 设置列宽——这是 Microsoft Word.js API 的限制，与协议或
> Add-In 实现无关。AI 在使用这套工具时，应先用 `update_table_format(columnWidths=...)`
> 设好列宽，再做合并。`update_table_format` 调用上**避开** `columnWidths` 仍可在合并
> 表上工作（用于 borders / alignment / cellPadding 等其它字段）。

**预期脚本输出尾部**：
```
  ✅ 表格数 = 1
  ✅ 首行已合并 (merge_cells)
  ✅ 首行单元格背景色 = #1F4E79 (update_table_cell)
  ✅ 第 2 行文本 = ['甲方', 'ACME Corp'] (update_table_row_column)
  ✅ tblGrid 列数 = 4，宽度 (dxa) = [...] (update_table_format)
==========================================================
✅ E2E 表格流水线通过；请打开下面文件目测视觉效果：
   manual_tests/.test_working/empty_<timestamp>.docx
```

**目测验收项**（打开脚本输出的 `.docx`）：

- [ ] **A.1** `word_merge_cells` 视觉：首行合并为单个跨整行单元格（Issue #8 项 2）
- [ ] **A.2** `word_update_table_cell` 视觉：首行 = 蓝底 (`#1F4E79`) + 居中 + 加粗 + 白字「甲方信息」（Issue #8 项 3）
- [ ] **A.3** `word_update_table_row_column` 视觉：第 2~5 行有「甲方/地址/联系人/日期」 + 对应值（Issue #8 项 4）
- [ ] **A.4** `word_update_table_format` 视觉：所有单元格有边框、列宽明显是 [120, 80, 80, 80] pt（首列宽于其他三列）、整表居中、单元格四周内边距均匀（Issue #8 项 5）
- [ ] **A.5** 标签列（甲方/地址/...）显示为灰底加粗
- [ ] **A.6** 缺省 `tableId` 命中：脚本第二阶段日志含 `omitted-tableId update_table_cell OK`（Issue #8 项 6）

---

## B. 错误码场景

### B.2 + B.3：已自动化 (随 `--mode tables` 一起跑)

`--mode tables` 在 OOXML 验证之后会自动跑 B.2 + B.3a + B.3b。

**符号约定**：✅/❌/⚠️ 表达**用例判定**，不表达正负向——负例按预期失败且错误码正确也是 ✅；
❌ 用例未通过；⚠️ 部分通过、有异常但可容忍（须挂账跟踪 issue）。负例的原始操作日志
（wire 层真实失败）以 📋 中性打印，不占用判定符号。输出形如：

```
🚨 B.2 错误码 + B.3a 冲突负例 + B.3b 混合列宽 merge 回归（自动触发）...
  --- B.2 不存在的 tableId → 3010 ELEMENT_NOT_FOUND（负例） ---
📋 更新单元格失败（负例预期内）: 3010: ...
    ✅ B.2 负例通过，错误码 3010: ...
  --- B.3a 相交合并（真·冲突，负例）→ 3014（暂容忍 3000，editor4ai#88） ---
📋 合并失败（负例预期内）: 3000: ...
    ⚠️  B.3a 实收 3000（已知缺口 editor4ai#88 接线后应 3014）: ...
  --- B.3b 混合列宽表不相交合并 (1,0)-(1,2) → 成功（editor4ai#85 回归） ---
    ✅ B.3b 混合列宽表 merge 成功（.merge() 路径可用）
```

- [ ] **B.2** 不存在的 `tableId` → 3010 ELEMENT_NOT_FOUND（Issue #8 项 8）
- [ ] **B.3a** 相交合并（首行已合并 + 再发 (0,0)-(1,2)）→ **失败**＝真·合并冲突
  （Word 拒绝：中文报文「尚未选定要合并的多个单元格」，officeCode=GeneralException）。
  OASP 归宿 3014 ALREADY_MERGED；Add-In 接线由 editor4ai#88 跟踪，接线前暂容忍 3000
  （⚠️ 输出计通过），接线后收紧为仅 3014（Issue #8 项 9 的现实版）
- [ ] **B.3b** 混合列宽表（首行已合并）上**不相交**合并 (1,0)-(1,2) → **成功**
  （office-editor4ai#85 修复回归：mergeCells 不再访问 `table.columns`；修复前此场景
  在执行前即抛裸 3000「混合的单元格宽度」）。
  注意：B.3b 成功后第 2 行前三格会合并（「甲方 ACME Corp」并入一格），目测 A.1~A.5
  请以 B.3b 之前的状态为准（或忽略第 2 行形变）。

### B.1 缺省 tableId + 光标不在表格内 → 3013 (Issue #8 项 7)

需要单独跑——脚本不能操控 Word 光标，必须人手把光标移出表格。

1. **不要关闭** `--mode tables` 留下来的 `.docx` 文件（仍开在 Word 里）
2. 在 Word 中**点击表格之外的某个普通段落**（让光标离开表格）
3. 重新启动 Add-In dev server（如果停了）
4. 跑：

```bash
uv run python manual_tests/word/test_word_table_e2e.py --mode b1
```

预期输出：
```
✅ B.1 错误码包含 3013: ...Cursor is not inside a table...
```

- [ ] **B.1** 错误码包含 `3013`

---

## C. 跨仓 / 业务闭环

### C.1 `word:get:documentStructure` 升级 (Issue #8 项 10)

> ⚠️ **本仓库 Server 侧未做该升级**。`docs/word.py::DocumentStructure` 仍只暴露
> `paragraphCount` / `tableCount` / `imageCount` / `sectionCount`。
> Add-In #23 (PR #27) 已经在响应里返回 `tables: [...]`，但 Server DTO 暂未声明，
> 字段会被 Pydantic 静默忽略。
>
> **可选跟进 Issue**（不阻塞本批 4 个 Tool 转 Stable）：
> 在 `WordGetDocumentStructureResponse.data` 上扩展 `tables: list[TableSummary] | None`。
> `TableSummary` DTO 已经在 `dtos/word.py` 中定义，仅缺一处装配。

- [ ] **C.1** （可选）记录新 Issue 跟踪 Server 侧 `documentStructure.tables[]` 装配

### C.2 AI 业务闭环（合同表头场景）(Issue #8 项 11)

让一个真实接入了 4 个 MCP Tool 的 LLM Agent 在空白文档执行：

> 「创建一份甲方信息表，首行作为横跨整行的'甲方信息'蓝底白字居中表头，下方为
> 标签-填写区两列结构（标签列灰底加粗、填写区白底），整表居中、列宽 [80, 200]、
> 内细外粗边框」

- [ ] **C.2** AI 单次会话内通过组合调用 4 个 Tool 完成；用户**无需右键合并 / 无需手工调样式**

---

## 验收完成后

1. 全部 A/B 项打勾 → 把这份清单的勾选状态复制到 Issue #8 评论
2. JIRA OF4AI-10 / OF4AI-11 添加 Server 侧完成评论（草稿见上一轮对话）
3. 创建 PR：`feat(word): expose 4 table OASP /word Draft tools (closes #8)`
4. PR 合并后关闭 Issue #8 + 转 OF4AI-10/11 为「完成」

---

**最后更新**：2026-04-30
**维护者**：JQQ &lt;jqq1716@gmail.com&gt;
