# 模板复用指南 | Template reuse guide

从**模板**生成新文件是 create-office-file 的头牌能力。两步：**实例化**（模板→成品）+ **填充**
（把占位/单元格/版式换成真实内容）。

## 0. 沙箱禁网 —— 模板必须是本地路径（关键约束）

运行环境**默认禁网**。传给 `office_run_script` 的 `template_uri`、以及脚本里读的模板路径，
**必须是 Computer 本地文件系统的路径**（如 `/home/user/templates/brand.dotx`），**不能是**
`http(s)://` / `oss://` 等在线地址——脚本里发起网络请求会被沙箱拦截。

若用户提供的是**在线模板 URL**：
1. 先用 **Computer 提供的网络/下载工具**把模板下载到 Computer 本地文件系统；
2. 再把**本地路径**作为 `template_uri` 传给 `office_run_script`。

## 1. 实例化：`instantiate_from_template(template, output)`

按模板扩展名自动分派（helper 内部处理，你不必关心细节）：

| 模板 | 成品 | 机制 | 依赖 |
|------|------|------|------|
| `.potx` | `.pptx` | content-type 互换（纯 zip 改写）| 无（CI 友好）|
| `.dotx` | `.docx` | LibreOffice 无头实例化 | **需装 soffice** |
| `.xltx` | `.xlsx` | LibreOffice 无头实例化 | **需装 soffice** |

```python
from office4ai.office.authoring.helpers import instantiate_from_template
instantiate_from_template("/local/brand.potx", "deck.pptx")   # 无需 soffice
instantiate_from_template("/local/brand.dotx", "report.docx") # 需 soffice
```

> soffice 缺失时 `.dotx`/`.xltx` 会抛 `SofficeNotFoundError`。`office_run_script` 启动日志会报告
> soffice 是否可用；若不可用，优先用 `.potx` 或从零生成。

## 2. 填充：按平台用对应 helper

实例化出成品后，用 helper 把「占位/结构」替换成真实内容——这些原语专为**保真复用**设计
（保留模板的样式/母版/图表，只换数据）：

| 平台 | helper | 用途 |
|------|--------|------|
| Word | `fill_sdt_controls(path, {alias: value})` | 按 alias 填 SDT 内容控件（占位符）|
| Word | `locate_by_style(path, style)` | 按段落样式定位锚点 |
| Excel | `fill_cells_lxml(path, {addr: value})` | 保图表改单元格 + 删公式缓存逼重算 |
| Excel | `locate_by_named_range(path, name)` | 定位命名区域 |
| PPT | `create_custom_layout(prs, master, name)` | 往母版新建自定义版式 |
| PPT | `locate_by_master(prs, name)` | 按名定位版式 |

```python
from office4ai.office.authoring.helpers import fill_sdt_controls
fill_sdt_controls("report.docx", {"文档标题": "2026 年度报告", "副标题": "XX 部门"})
```

完整入口见 `from office4ai.office.authoring.helpers import ...`；读 `scripts/from_template.py` 学串联。

## 3. 何时从零、何时从模板

- **有模板** → 优先从模板（保企业品牌/版式/图表，改数据即可）。
- **无模板** → 从零生成（`scripts/gen_word.py` / `gen_ppt.py` / `gen_excel.py`）。
- 需要**把参考文件脱胎成模板** → 另见 `extract-template` SKILL（W3）。
