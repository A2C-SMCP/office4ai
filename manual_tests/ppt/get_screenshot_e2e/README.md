# PPT Get Screenshot E2E Tests

`ppt:get:slideScreenshot` 事件的端到端测试。

## 测试场景

| # | 名称 | Fixture | 描述 |
|---|------|---------|------|
| 1 | PNG 截图 | colored_slide.pptx | 以默认 PNG 格式获取彩色幻灯片截图 |
| 2 | JPEG 截图 | colored_slide.pptx | 以 JPEG 格式获取截图 |
| 3 | Base64 格式验证 | colored_slide.pptx | 验证截图返回数据是合法的 Base64 编码 |

## 运行

```bash
uv run python manual_tests/ppt/get_screenshot_e2e/test_screenshot.py --test 1
uv run python manual_tests/ppt/get_screenshot_e2e/test_screenshot.py --test all
uv run python manual_tests/ppt/get_screenshot_e2e/test_screenshot.py --list
```

## CLI 选项

| 选项 | 说明 |
|------|------|
| `--test N \| all` | 运行第 N 个测试或全部 |
| `--list` | 列出所有测试用例 |
| `--no-auto-open` | 手动打开文档模式 |
| `--always-cleanup` | 失败时也清理测试文件 |

## 错误码 3007 / 3016 为何没有真机用例（issue #90 / oasp#23）

oasp#23 给本事件补了 `3007 FORMAT_NOT_SUPPORTED`（`options.format` 枚举合法但目标渲染不出）
与 `3016 API_NOT_SUPPORTED`（整个截图能力在当前宿主不可用）两行。**两者在真机上都不可构造**：

- `3007` —— 需要一个「枚举内但本宿主产不出」的 `format` 取值，而本目录用例已把可用取值验证为
  **成功**路径；没有这样的取值可请求。
- `3016` —— 需要一个**不支持截图**的宿主，而真机 E2E 恰恰跑在支持截图的宿主上。

硬写只会得到永远失败的用例（xfail 三态**不赦免**「本应失败却成功」，见
`manual_tests/error_case.py`）。故两码由**契约层**覆盖（mock 可模拟宿主拒绝）：
`tests/contract_tests/ppt/test_get_slide_screenshot.py::test_get_slide_screenshot_format_not_supported_3007`
与 `..._api_not_supported_3016`。

同构说明见 `manual_tests/word/export_content_e2e/README.md`（`word:export:content` 同因）。
