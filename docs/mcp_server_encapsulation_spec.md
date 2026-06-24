# MCP Server 封装规范

> **Office4AI MCP Server 架构设计与实现规范**
> **状态**: 已确认 | **日期**: 2026-02-08

---

## 1. 设计定位

OfficeMCPServer 是 **Office 级别**的统一 MCP Server，一个 Server 实例同时处理 Word、PPT、Excel 三种文档类型。Socket.IO Server 作为纯工具通道，不包含安全检查或租户隔离，鉴权与权限完全依赖 MCP 协议与产品层约束。

### 核心原则

| 原则 | 描述 |
|------|------|
| **纯工具定位** | Socket.IO 是能力提供通道，不做安全检查，不做多租户隔离 |
| **生命周期绑定** | Socket.IO Server 的生命周期与 MCP Server 一致 |
| **精简面向 AI** | MCP 工具数量精简、语义清晰、易于 AI 理解 |
| **对齐 ide4ai** | 架构模式参考 ide4ai，保持两个项目的结构一致性 |

---

## 2. 架构总览

```
AI Agent / Claude
    ↓ (MCP Protocol: stdio / SSE / streamable-http)
OfficeMCPServer (统一入口, server_name="office4ai")
    ├─ list_tools()  → 返回所有平台工具 (word_*, ppt_*, excel_*)
    ├─ call_tool()   → 路由到对应 Tool.execute()
    └─ list_resources() / read_resource() → (本次不实现)
    ↓
BaseTool.execute() (声明式子类, 通用执行逻辑)
    ├─ 验证输入 (Pydantic InputModel)
    ├─ 构建 OfficeAction(category, action_name, params)
    └─ 调用 workspace.execute(action)
    ↓
OfficeWorkspace (单实例, 管理所有平台连接)
    ├─ execute() → wrap_request() → sio.call()
    ├─ ConnectionManager (全局单例, 多 namespace 共存)
    └─ Socket.IO Server (HTTP:3000 / HTTPS:4443)
    ↓
Office Add-In Clients (/word, /ppt, /excel namespaces)
```

---

## 3. 生命周期设计

### 3.1 Async 初始化方案：分离构造与启动

解决 Python `__init__` 不支持 async 的经典问题。采用与现有 OfficeWorkspace 的 `start()/stop()` 生命周期天然一致的方案。

```python
class OfficeMCPServer(BaseMCPServer):
    def __init__(self, config: MCPServerConfig) -> None:
        # 同步：创建 workspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
        )
        # 同步：注册 Tools (持有 workspace 引用, 此时 workspace 未 start)
        super().__init__(config, server_name="office4ai")

    async def _async_startup(self) -> None:
        """MCP run() 时调用, 启动 async 服务"""
        await self.workspace.start()

    async def _async_shutdown(self) -> None:
        """MCP 关闭时调用, 停止 async 服务"""
        await self.workspace.stop()

    async def run(self) -> None:
        """Override: 先启动 workspace, 再进入 MCP 事件循环"""
        await self._async_startup()
        try:
            await super().run()
        finally:
            await self._async_shutdown()
```

### 3.2 启动顺序

```
1. __init__()
   ├─ 创建 OfficeWorkspace 实例 (同步, 未启动)
   ├─ _register_tools()  → Tool 持有 workspace 引用
   ├─ _register_resources()  → (空实现)
   └─ _setup_handlers()  → MCP 协议处理器

2. run()
   ├─ await _async_startup()
   │   └─ await workspace.start()  → Socket.IO Server 就绪
   ├─ await super().run()  → 进入 MCP 事件循环 (stdio/SSE/HTTP)
   └─ finally: await _async_shutdown()
       └─ await workspace.stop()  → 清理 Socket.IO Server
```

---

## 4. 工具依赖注入

### 4.1 注入方式：直接注入 OfficeWorkspace

Tool 直接持有 OfficeWorkspace 实例引用，不引入额外的中间层抽象。

```python
class BaseTool(ABC):
    def __init__(self, workspace: OfficeWorkspace) -> None:
        self.workspace = workspace
```

**理由**：office4ai 没有 IDE 这层抽象（没有 terminal），直接注入 workspace 简单直接。将来需要扩展时再引入中间层。

### 4.2 document_uri 参数设计

每个工具调用**必须传入 `document_uri`** 参数。显式优于隐式。

```python
class WordInsertTextInput(BaseModel):
    document_uri: str = Field(..., description="目标文档的 URI")
    text: str = Field(..., description="要插入的文本")
    format: TextFormat | None = Field(None, description="文本格式")
```

AI Agent 通过 ConnectionManager 的信息（或工具错误信息）获知可用文档。

---

## 5. 声明式 BaseTool 设计

### 5.1 核心模式

BaseTool 封装通用 `execute()` 逻辑，子类只需声明元数据。采用 Template Method 模式，提供 `format_result()` hook 供子类定制返回格式。

```python
class BaseTool(ABC):
    """声明式工具基类"""

    def __init__(self, workspace: OfficeWorkspace) -> None:
        self.workspace = workspace

    # ── 子类必须声明的元数据 ──

    @property
    @abstractmethod
    def name(self) -> str:
        """工具名称, 如 'word_insert_text'"""
        raise NotImplementedError

    @property
    @abstractmethod
    def description(self) -> str:
        """工具描述, 面向 AI 的自然语言说明"""
        raise NotImplementedError

    @property
    @abstractmethod
    def input_schema(self) -> dict[str, Any]:
        """JSON Schema, 通常由 InputModel.model_json_schema() 生成"""
        raise NotImplementedError

    @property
    @abstractmethod
    def category(self) -> str:
        """平台类别: 'word' | 'ppt' | 'excel'"""
        raise NotImplementedError

    @property
    @abstractmethod
    def event_name(self) -> str:
        """Socket.IO 事件名, 如 'insert:text'"""
        raise NotImplementedError

    @property
    @abstractmethod
    def input_model(self) -> type[BaseModel]:
        """Pydantic InputModel 类"""
        raise NotImplementedError

    # ── 通用执行逻辑 ──

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        通用执行流程:
        1. 验证输入
        2. 提取 document_uri 和业务参数
        3. 构建 OfficeAction
        4. 调用 workspace.execute()
        5. 格式化返回
        """
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, self.input_model)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 提取参数
        params = validated.model_dump(exclude_none=True)
        document_uri = params.pop("document_uri")

        # 3. 构建 OfficeAction
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": document_uri, **params},
        )

        # 4. 执行
        try:
            obs = await self.workspace.execute(action)
        except Exception as e:
            return {"success": False, "error": str(e)}

        # 5. 格式化返回 (hook)
        return self.format_result(obs)

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """
        默认返回格式化 hook. 子类可 override.
        默认行为: 返回 JSON 结构.
        获取类工具可 override 返回纯文本/Markdown.
        """
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        return {"success": True, "data": obs.data}

    def validate_input(self, arguments: dict, model: type[T]) -> T:
        """Pydantic 输入验证"""
        return model.model_validate(arguments)
```

### 5.2 子类示例

```python
class WordInsertTextTool(BaseTool):
    """Word 插入文本工具 - 声明式实现"""

    @property
    def name(self) -> str:
        return "word_insert_text"

    @property
    def description(self) -> str:
        return "在 Word 文档的当前光标位置插入文本。可选择指定文本格式（字体、大小、加粗等）。"

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertTextInput.model_json_schema()

    @property
    def category(self) -> str:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertTextInput
```

```python
class WordGetVisibleContentTool(BaseTool):
    """获取可见内容工具 - override format_result 返回纯文本"""

    @property
    def name(self) -> str:
        return "word_get_visible_content"

    # ... 其他元数据 ...

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回可读文本而非 JSON"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        # 将文档内容格式化为 Markdown 或纯文本
        content = obs.data.get("content", "")
        return {"success": True, "content": content}
```

---

## 6. 工具命名规范

### 6.1 命名模式: `<platform>_<action>_<resource>`

```
word_insert_text        # Word 插入文本
word_append_text        # Word 追加文本
word_replace_text       # Word 查找替换
word_get_selected_content  # Word 获取选中内容
ppt_insert_text         # PPT 插入文本 (未来)
excel_set_cell_value    # Excel 设置单元格 (未来)
```

**理由**: MCP Server 是 Office 级别的统一服务，平台前缀是必要的区分标识。

### 6.2 Socket.IO 事件映射

| MCP Tool Name | Socket.IO Event | category | event_name |
|---------------|----------------|----------|------------|
| `word_insert_text` | `word:insert:text` | `word` | `insert:text` |
| `word_append_text` | `word:append:text` | `word` | `append:text` |
| `word_replace_text` | `word:replace:text` | `word` | `replace:text` |

---

## 7. Word MCP 工具列表 (MVP)

### 7.1 精简决策记录

| 原始工具 | 决策 | 理由 |
|---------|------|------|
| `word_get_selected_content` | **保留** | 配合用户选区使用，核心工具 |
| `word_get_visible_content` | **保留** | 获取上下文的合适粒度 |
| `word_get_document_structure` | **移到 Resource** | 元信息，非操作。留到 A2C 阶段 |
| `word_get_document_stats` | **移到 Resource** | 元信息，非操作。留到 A2C 阶段 |
| `word_get_styles` | **移到 Resource** | 参考信息，非操作。留到 A2C 阶段 |
| `word_insert_text` | **保留** | 核心编辑工具 |
| `word_append_text` | **保留** | 语义清晰，面向末尾追加场景 |
| `word_replace_selection` | **删除** | `replace_text` 的 find & replace 已覆盖 |
| `word_replace_text` | **保留** | find & replace，依赖 Office JS Search API |
| `word_select_text` | **延后** | MVP 不实现，replace_text + insert 已覆盖主要场景 |
| `word_insert_image` | **保留** | 多媒体插入，独立 schema |
| `word_insert_table` | **保留** | 多媒体插入，独立 schema |
| `word_insert_equation` | **保留** | 多媒体插入，独立 schema |
| `word_insert_toc` | **保留** | 高级功能，几乎无参 |
| `word_export_content` | **延后** | MVP 不实现，不常用 |

### 7.2 MVP 工具清单 (9 个)

| # | MCP Tool Name | 类型 | Socket.IO Event | 返回格式 |
|---|--------------|------|----------------|---------|
| 1 | `word_get_selected_content` | 获取 | `word:get:selectedContent` | 适配文本 |
| 2 | `word_get_visible_content` | 获取 | `word:get:visibleContent` | 适配文本 |
| 3 | `word_insert_text` | 操作 | `word:insert:text` | JSON |
| 4 | `word_append_text` | 操作 | `word:append:text` | JSON |
| 5 | `word_replace_text` | 操作 | `word:replace:text` | JSON |
| 6 | `word_insert_image` | 多媒体 | `word:insert:image` | JSON |
| 7 | `word_insert_table` | 多媒体 | `word:insert:table` | JSON |
| 8 | `word_insert_equation` | 多媒体 | `word:insert:equation` | JSON |
| 9 | `word_insert_toc` | 高级 | `word:insert:toc` | JSON |

### 7.3 Replace 语义

`word_replace_text` 完全依赖 Office JS Search API 实现查找替换，等价于 Word 的 Ctrl+H。Python 侧只传参数，不做文本拼接/位置计算。

---

## 8. 错误处理

### 8.1 统一 error string

所有错误统一返回 `{"success": false, "error": "错误描述"}`，由 AI 通过自然语言理解错误原因。不区分错误类型，不使用 error_code。

```python
# 文档未连接
{"success": false, "error": "Document not connected: file:///path/to/doc.docx"}

# 参数验证失败
{"success": false, "error": "Validation error: 'text' field is required"}

# 执行超时
{"success": false, "error": "Operation timed out after 10 seconds"}
```

### 8.2 与 ide4ai 对齐

与 ide4ai 的 Tool 返回模式一致：`{"success": bool, "error": str | None, ...}`。

---

## 9. 文件结构

```
office4ai/
├── a2c_smcp/
│   ├── server.py              # BaseMCPServer (修改: 添加 async lifecycle hooks)
│   ├── config.py              # MCPServerConfig (可能需要添加 socketio_port)
│   ├── tools/
│   │   ├── base.py            # BaseTool (重写: 声明式模式 + format_result hook)
│   │   └── word/
│   │       ├── __init__.py
│   │       ├── get_selected_content.py   # WordGetSelectedContentTool
│   │       ├── get_visible_content.py    # WordGetVisibleContentTool
│   │       ├── insert_text.py            # WordInsertTextTool
│   │       ├── append_text.py            # WordAppendTextTool
│   │       ├── replace_text.py           # WordReplaceTextTool
│   │       ├── insert_image.py           # WordInsertImageTool
│   │       ├── insert_table.py           # WordInsertTableTool
│   │       ├── insert_equation.py        # WordInsertEquationTool
│   │       └── insert_toc.py             # WordInsertTOCTool
│   └── resources/
│       └── base.py            # BaseResource (不变)
├── office/
│   └── mcp/
│       └── server.py          # OfficeMCPServer (重写: async lifecycle + tool 注册)
└── environment/
    └── workspace/             # (不变, 已实现)
        ├── office_workspace.py
        └── socketio/
```

---

## 10. OfficeMCPServer 实现

```python
class OfficeMCPServer(BaseMCPServer):
    """Office 级别的统一 MCP Server"""

    def __init__(self, config: MCPServerConfig) -> None:
        # 创建 OfficeWorkspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
        )
        super().__init__(config, server_name="office4ai")

    def _register_tools(self) -> None:
        """注册所有平台的工具"""
        # Word 工具
        from office4ai.a2c_smcp.tools.word import (
            WordGetSelectedContentTool,
            WordGetVisibleContentTool,
            WordInsertTextTool,
            WordAppendTextTool,
            WordReplaceTextTool,
            WordInsertImageTool,
            WordInsertTableTool,
            WordInsertEquationTool,
            WordInsertTOCTool,
        )

        word_tools = [
            WordGetSelectedContentTool(self.workspace),
            WordGetVisibleContentTool(self.workspace),
            WordInsertTextTool(self.workspace),
            WordAppendTextTool(self.workspace),
            WordReplaceTextTool(self.workspace),
            WordInsertImageTool(self.workspace),
            WordInsertTableTool(self.workspace),
            WordInsertEquationTool(self.workspace),
            WordInsertTOCTool(self.workspace),
        ]

        for tool in word_tools:
            self.tools[tool.name] = tool

        # PPT 工具 (未来)
        # Excel 工具 (未来)

    def _register_resources(self) -> None:
        """本次不实现任何 Resource"""
        pass

    async def _async_startup(self) -> None:
        await self.workspace.start()

    async def _async_shutdown(self) -> None:
        await self.workspace.stop()

    async def run(self) -> None:
        await self._async_startup()
        try:
            await super().run()
        finally:
            await self._async_shutdown()
```

---

## 11. BaseMCPServer 修改

需要在 `BaseMCPServer` 中添加 async lifecycle 钩子，使其可被子类 override：

```python
class BaseMCPServer(ABC):
    # ... 现有代码不变 ...

    async def _async_startup(self) -> None:
        """async 启动钩子, 子类可 override"""
        pass

    async def _async_shutdown(self) -> None:
        """async 关闭钩子, 子类可 override"""
        pass

    async def run(self) -> None:
        """修改: 支持 startup/shutdown 钩子"""
        await self._async_startup()
        try:
            transport = self.config.transport
            if transport == "stdio":
                await self._run_stdio()
            elif transport == "sse":
                await self._run_sse()
            elif transport == "streamable-http":
                await self._run_streamable_http()
            else:
                raise ValueError(f"Unsupported transport: {transport}")
        finally:
            await self._async_shutdown()
```

---

## 12. 测试策略

### 12.1 单元测试结构

```
tests/unit_tests/
└── a2c_smcp/
    └── tools/
        └── word/
            ├── test_insert_text.py
            ├── test_append_text.py
            ├── test_replace_text.py
            ├── test_get_selected_content.py
            ├── test_get_visible_content.py
            ├── test_insert_image.py
            ├── test_insert_table.py
            ├── test_insert_equation.py
            └── test_insert_toc.py
```

### 12.2 测试策略

- **Mock OfficeWorkspace**: 每个 Tool 测试 mock `workspace.execute()` 的返回值
- **验证 OfficeAction 构建**: 确认 category, event_name, params 的正确性
- **验证输入校验**: 测试缺少必要参数、参数类型错误等边界情况
- **验证 format_result**: 确认获取类工具返回适配文本，操作类工具返回 JSON

---

## 13. 实现优先级

| 阶段 | 内容 | 依赖 |
|------|------|------|
| **P0** | 修改 BaseMCPServer (async lifecycle hooks) | 无 |
| **P0** | 重写 BaseTool (声明式模式) | 无 |
| **P1** | 实现 OfficeMCPServer (workspace 集成 + tool 注册) | P0 |
| **P1** | 实现 9 个 Word Tool 类 | P0 |
| **P2** | 单元测试 | P1 |
| **P3** | E2E 测试 (需要 Add-In 配合) | P1 |

---

## 附录 A: 关键决策汇总

| # | 决策项 | 选择 | 备选方案 |
|---|--------|------|---------|
| 1 | MCP Server 级别 | Office 级别 (统一) | 每平台一个 Server |
| 2 | 多租户隔离 | 不实现，纯工具 | Session 隔离 |
| 3 | Replace 语义 | 依赖 Office JS Search API | Python 侧智能匹配 |
| 4 | insert/append 关系 | 按操作语义拆分，保留两个 | 合并为一个 insert |
| 5 | replace_selection | 删除，replace_text 覆盖 | 保留 |
| 6 | Tool 依赖注入 | 直接注入 OfficeWorkspace | IDE 中间层 |
| 7 | Async 初始化 | 分离构造与启动 | async classmethod 工厂 |
| 8 | document_uri | 每次必传 | 活跃文档概念 |
| 9 | 获取类工具 | 保留独立 (get_selected/visible) | 合并为 read_document |
| 10 | get_styles/structure/stats | 移到 Resource (A2C 阶段) | 保留为 Tool |
| 11 | 多媒体插入工具 | 保持独立工具 | 合并为 insert_object |
| 12 | select_text | MVP 延后 | 保留 |
| 13 | export_content | MVP 延后 | 保留 |
| 14 | Resource | 本次不实现 | 实现基础 Resource |
| 15 | 工具命名 | 平台前缀下划线 (word_insert_text) | 命名空间风格 |
| 16 | Tool 实现架构 | 每平台独立 Tool 类 | 通用 Tool + 配置 |
| 17 | BaseTool 模式 | 声明式子类 + hook | 每个工具完整实现 |
| 18 | 返回格式 | 根据工具类型适配 (hook) | 统一 JSON |
| 19 | 错误处理 | 统一 error string | 结构化错误码 |
| 20 | Workspace 粒度 | 一个 OfficeWorkspace 管全部 | 每平台独立 |
| 21 | Tool 文件位置 | a2c_smcp/tools/word/ | office/mcp/tools/ |
| 22 | Excel 写操作返回 | 最小返回 + write-then-read (Plan A) | Server 侧补读回填 (Plan B) |

> 决策 22 详见 [`docs/discussions/excel-minimal-return-decision.md`](discussions/excel-minimal-return-decision.md)（milestone #3 / issue #26）。约定登记于基类 `BaseTool.format_result()` docstring（写工具继承的单一源）。

---

## 附录 B: 与 ide4ai 架构对齐对照

| 概念 | ide4ai | office4ai |
|------|--------|-----------|
| 顶层抽象 | IDE (gym.Env) | OfficeWorkspace |
| MCP Server | PythonIDEMCPServer | OfficeMCPServer |
| Tool 基类 | BaseTool(ide) | BaseTool(workspace) |
| Tool 实现 | 6 个通用工具 | 9 个 Word 工具 (MVP) |
| Resource | WindowResource | 不实现 (A2C 阶段) |
| 通信层 | LSP (进程内) | Socket.IO (网络) |
| Workspace | PyWorkspace (LSP) | OfficeWorkspace (Socket.IO) |
| 配置 | MCPServerConfig | MCPServerConfig (共用) |
| 传输 | stdio/SSE/HTTP | stdio/SSE/HTTP |
