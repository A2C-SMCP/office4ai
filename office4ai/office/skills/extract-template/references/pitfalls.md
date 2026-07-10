# 适配坑 | Pitfalls & gotchas（抽取参考文件为模板）

## 沙箱约束

- **原地 vs 读入-写出**：`wrap_in_sdt` 是**原地编辑**——先 `shutil.copy(参考文件, "本地名")` 进工作目录再抽；
  `replace_runs_with_token` / PPT / Excel 是**读入-写出**，直接指定工作目录内的输出名，不必先 copy。
- **默认禁网**：参考文件必须先落到 Computer 本地，用本地路径传 `template_uri`（进 FS 读白名单）。
- **import 白名单**：`docx` / `pptx` / `openpyxl` / `lxml` / `docxtpl` / `PIL` + 安全 stdlib（含 `shutil`）+
  `office4ai.office.authoring.helpers`。

## Word · SDT 抽取（`wrap_in_sdt`）

- `para_index` 用 `locate_by_style` 返回的 `.index`，**别自己数段落**——两者同坐标系。越界抛 `AnchorNotFoundError`。
- `alias` 就是回填键：`wrap_in_sdt(alias="标题")` ↔ `fill_sdt_controls({"标题": 值})`。多个同名 alias 时，
  回填用**有序序列**而非字典（字典同名只填第一个）。
- `placeholder=None` 保留段落原文本作占位；给了 `placeholder` 则把 run 归一为单个占位 run（顺带消解 run-split）。
- 逐段 wrap 是安全的：wrap 不改变正文 `w:p` 的文档顺序/数量，故先 `locate_by_style` 一次拿到的 index 列表，
  顺序 wrap 各段互不影响。

## Word · token 抽取（`replace_runs_with_token`）

- `{{token}}` 会落进**单个 run**（helper 已跨 run 合并），docxtpl 才认得。别手工在模板里敲占位——Word 自动更正
  可能把它拆断。
- `strict=True`（默认）：任一实值一次都没命中即抛 `AnchorNotFoundError`（fail-fast，防抽错文件）；容忍缺失用
  `strict=False`。注意与 `fill_cells_lxml`（默认 `strict=False`）不对称。
- 实值要**足够独特**：替换按段内拼接全文子串匹配，太短的实值（如「1」）可能误伤。选可辨识的整段值。

## PPT · 版式抽取（`promote_slide_to_layout`）

- 样板页**不要带图片/超链接等外部媒体**：形状 XML 深拷贝进新版式，但媒体/hlink 的 `r:id` 只在源 slide 部件里，
  会在新版式部件悬空。`promote_slide_to_layout` 命中此情况会**快速失败**抛 `ExternalMediaError`（而非静默产出损坏
  文件）；请选不含媒体的样板页，或先移除媒体。
- `remove_source=True` 抽完删源页得干净模板；`False` 则保留源页作样例。删源页后 `slide_index` 之后的序号会前移，
  一次只抽一页最省心。
- `base_layout_index` 默认 6（空白版式）作合法外壳；越界抛 `ValueError`。

## Excel · 命名区域抽取

- 保图表是头等事：清示例值用 `fill_cells_lxml(..., {sheet: {cell: None}})`，**别** `openpyxl` 整体载入-重存
  （对复杂图表/条件格式建模不完整，有丢失风险）。
- `locate_by_named_range` 找不到返回**空列表**（不抛异常）；脚本自己判空并 `raise SystemExit(...)` fail-fast。
- `NamedRange.ref`（如 `$B$2:$C$30`）是尽力解析的区域串；要按区域展开成单元格清空，需自行解析区间，或直接
  用 `BLANK_CELLS` 显式列出示例单元格（参考脚本采用后者，更稳）。

## 通用

- **抽取是显式的**：你决定哪些是可变部分。先勘锚点（`locate_*` / 遍历）再抽，别指望自动探测。
- 抽完 `print(...)` 摘要便于从 `logs` 确认；`summary.produced` 列出落盘的模板产物。
- **闭环校验**：抽出的模板最好当场用对应回填手段跑一遍（SDT→fill、token→docxtpl、layout→add_slide），
  确认可复用——这也是本 SKILL 集成测试的做法。
