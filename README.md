# Office4AI

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-0.1.1rc1-green.svg)](https://github.com/JQQ/office4ai)

**Office4AI** 是一个 MCP Server，让 AI Agent 通过 Office Add-In 实时读写 Office 文档。

> **当前状态：** Word（9 个工具）已实现，PowerPoint 和 Excel 开发中。

## 支持平台

| 平台        | 状态   | 工具数 |
|-------------|--------|--------|
| Word        | 已就绪 | 9      |
| PowerPoint  | 规划中 | —      |
| Excel       | 规划中 | —      |

## 快速开始

### 1. 安装证书（仅首次）

Office4AI 使用 HTTPS 与 Office Add-In 通信。运行 `setup` 生成并安装本地 CA 证书：

```bash
uvx office4ai-mcp setup
```

将生成本地 CA 和服务器证书，并将 CA 安装到系统信任存储（需要管理员权限）。

### 2. 启动服务

```bash
uvx office4ai-mcp serve
```

### 3. 配置 MCP 客户端

#### Claude Code（推荐）

```bash
claude mcp add office4ai -- uvx --from office4ai office4ai-mcp
```

#### Claude Desktop / Cursor 等 MCP 客户端

在 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "office4ai": {
      "command": "uvx",
      "args": ["office4ai-mcp", "serve"]
    }
  }
}
```

### 4. 安装 Office Add-In

Office Add-In 通过 Socket.IO 连接 Microsoft Office 与 MCP Server。

> Add-In 安装说明将另行提供，敬请关注。

## 可用工具

### Word 工具（9 个）

| 工具 | 说明 |
|------|------|
| `word_get_selected_content` | 获取当前选中内容——文本、元素及元数据 |
| `word_get_visible_content` | 获取当前视口中可见的内容 |
| `word_insert_text` | 在光标处插入文本，支持格式（粗体、斜体、字体、颜色、样式） |
| `word_append_text` | 在文档开头或末尾追加文本 |
| `word_replace_text` | 查找并替换文本（等同 Ctrl+H），支持大小写敏感和全词匹配 |
| `word_insert_image` | 插入 base64 编码的图片，支持尺寸和替代文字 |
| `word_insert_table` | 插入表格，指定行列数、数据和样式 |
| `word_insert_equation` | 插入 LaTeX 公式（默认为行内公式） |
| `word_insert_toc` | 插入目录（可配置标题级别） |

所有工具均需要 `document_uri` 参数来标识目标文档。

## CLI 命令参考

```
office4ai-mcp <command>
```

| 命令      | 说明 |
|-----------|------|
| `serve`   | 启动 MCP Server（未指定命令时的默认行为） |
| `setup`   | 生成证书并将 CA 安装到系统信任存储 |
| `cleanup` | 从信任存储移除 CA 并删除证书文件 |

### 服务器选项

可通过 CLI 参数或环境变量设置：

| 选项 | 环境变量 | 默认值 | 说明 |
|------|---------|--------|------|
| `--transport` | `TRANSPORT` | `stdio` | MCP 传输方式：`stdio`、`sse` 或 `streamable-http` |
| `--host` | `HOST` | `127.0.0.1` | 服务器绑定地址 |
| `--port` | `PORT` | `8000` | MCP HTTP 端口（用于 `sse`/`streamable-http`） |
| `--socketio-port` | `SOCKETIO_PORT` | `3000` | Socket.IO 端口（Add-In 连接） |

### 证书位置

证书默认存储在 `~/.office4ai/certs/`，可通过 `OFFICE4AI_CERT_DIR` 环境变量覆盖。

### 日志位置

日志默认存储在平台标准目录，可通过 `OFFICE4AI_LOG_DIR` 环境变量覆盖：

| 平台    | 默认路径                                         |
|---------|--------------------------------------------------|
| macOS   | `~/Library/Logs/office4ai`                       |
| Linux   | `~/.local/state/office4ai/log`                   |
| Windows | `C:\Users\<user>\AppData\Local\office4ai\Logs`   |

设置 `OFFICE4AI_LOG_DIR=""` 可禁用文件日志。

## 系统要求

- Python 3.10+
- macOS 或 Windows
- 支持 Add-In 的 Microsoft Office（Word 桌面版或 Word Online）

## 开发

开发环境配置、常用命令、代码规范和贡献指南请参考 [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md)。

## 许可证

MIT 许可证 - 详见 [LICENSE](LICENSE)。

## 联系方式

- 作者：JQQ
- 邮箱：jqq1716@gmail.com
- GitHub：[@JQQ](https://github.com/JQQ)
