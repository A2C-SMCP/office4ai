# PPT Insert Shape E2E Tests

`ppt:insert:shape` 事件的端到端测试。

## 测试文件

### test_shape_insert.py — 基本形状插入

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 插入矩形 | empty.pptx | 插入 Rectangle 形状 |
| 2 | 插入圆形 | empty.pptx | 插入 Circle 形状 |
| 3 | 带文本形状 | empty.pptx | 插入带 text='Hello Shape' 的矩形 |
| 4 | 带样式形状 | empty.pptx | 插入带红色填充和蓝色边框的矩形 |
| 5 | 带字体形状 (0.4.0) | empty.pptx | RoundedRectangle 带 text+font(size24/微软雅黑/红/粗/双删除线)；**读回 python-pptx 校验** bold+size+doubleStrikethrough 真生效 |
| 6 | Line 拒绝 font (负例, 0.4.0) | empty.pptx | 无文本框的 Line 传 font/text → **期望 4002 静态拒绝**（text-capable 门控，验证 4002/3016 语义分离） |

> 用例 5/6 覆盖 OASP 0.4.0 `insert:shape` 的 `font`（[office4ai#76](https://github.com/A2C-SMCP/office4ai/issues/76)）。**前提**：Add-In 侧须为 0.4.0（否则握手 2006 断连）。用例 5 用 `double_strikethrough`（snake_case 入参 → camelCase wire）作为字体真落地的判据。

### test_shape_types.py — 批量形状类型

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | 批量基本形状 | empty.pptx | 连续插入 Rectangle, Circle, Oval, Triangle |
| 2 | 箭头线条 | empty.pptx | 插入 Arrow 和 Line |
| 3 | 星形文本框 | empty.pptx | 插入 Star, RoundedRectangle, TextBox |

## 运行

```bash
# 基本形状
uv run python manual_tests/ppt/insert_shape_e2e/test_shape_insert.py --test all

# 形状类型
uv run python manual_tests/ppt/insert_shape_e2e/test_shape_types.py --test all
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |
