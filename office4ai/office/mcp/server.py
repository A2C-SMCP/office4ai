# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

import asyncio
import os
import shutil
import sys
from pathlib import Path
from typing import Any

from loguru import logger

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.resources.per_file_window import (
    PerFileWindowResource,
    affected_window_uris,
    create_per_file_window,
    per_file_window_base_uri,
)
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.certs.paths import get_cert_dir
from office4ai.certs.trust_store import get_trust_store
from office4ai.certs.validator import CertStatus, get_cert_expiry_info, validate_certs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
    normalize_document_uri,
)


class OfficeMCPServer(BaseMCPServer):
    """
    Office 级别的统一 MCP Server | Office-level unified MCP Server

    一个 Server 实例同时处理 Word、PPT、Excel 三种文档类型。
    Socket.IO Server 的生命周期与 MCP Server 一致。
    """

    # authoring SKILL 包分发根环境变量（部署/测试可覆盖默认包内 office/skills/）。
    # Env override for the authoring SKILL package root (deploy/test override the packaged office/skills/).
    SKILLS_ROOT_ENV = "OFFICE4AI_SKILLS_ROOT"

    def __init__(self, config: MCPServerConfig, cert_dir: Path | None = None, skills_root: Path | None = None) -> None:
        # 同步：创建 workspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
            cert_dir=cert_dir,
        )
        # SKILL 包分发根须在 super().__init__() → _register_resources() 之前就位。
        self._skills_root = self._resolve_skills_root(skills_root)
        # W4b-2 (#65): fullscreen 归属台账——记录每个文件被 AI 操作的先后序号（单调递增），
        # 用于断连后「顺延次新活跃者」的裁决（契约③）。回调运行时更新，__init__ 先就位。
        self._activity_order: dict[str, int] = {}
        self._activity_seq: int = 0
        super().__init__(config=config, server_name="office4ai")

    @classmethod
    def _resolve_skills_root(cls, skills_root: Path | None) -> Path:
        """解析 SKILL 包分发根：显式参数 > 环境变量 > 包内默认 ``office/skills/``。"""
        if skills_root is not None:
            return Path(skills_root)
        env = os.environ.get(cls.SKILLS_ROOT_ENV)
        if env:
            return Path(env)
        return Path(__file__).resolve().parent.parent / "skills"

    def _register_tools(self) -> None:
        """注册所有平台的工具 | Register all platform tools"""
        from office4ai.a2c_smcp.tools.word import (
            WordAppendTextTool,
            WordDeleteCommentTool,
            WordExportContentTool,
            WordGetCommentsTool,
            WordGetDocumentStatsTool,
            WordGetDocumentStructureTool,
            WordGetSelectedContentTool,
            WordGetSelectionTool,
            WordGetStylesTool,
            WordGetVisibleContentTool,
            WordInsertCommentTool,
            WordInsertEquationTool,
            WordInsertImageTool,
            WordInsertTableTool,
            WordInsertTextTool,
            WordInsertTOCTool,
            WordMergeCellsTool,
            WordReplaceSelectionTool,
            WordReplaceTextTool,
            WordReplyCommentTool,
            WordResolveCommentTool,
            WordRunScriptTool,
            WordSelectTextTool,
            WordUpdateTableCellTool,
            WordUpdateTableFormatTool,
            WordUpdateTableRowColumnTool,
        )

        word_tools = [
            # Get tools
            WordGetSelectedContentTool(self.workspace),
            WordGetVisibleContentTool(self.workspace),
            WordGetSelectionTool(self.workspace),
            WordGetDocumentStructureTool(self.workspace),
            WordGetDocumentStatsTool(self.workspace),
            WordGetStylesTool(self.workspace),
            # Text operation tools
            WordInsertTextTool(self.workspace),
            WordAppendTextTool(self.workspace),
            WordReplaceTextTool(self.workspace),
            WordReplaceSelectionTool(self.workspace),
            WordSelectTextTool(self.workspace),
            # Multimedia tools
            WordInsertImageTool(self.workspace),
            WordInsertTableTool(self.workspace),
            WordInsertEquationTool(self.workspace),
            WordInsertTOCTool(self.workspace),
            # Table operation tools (OASP /word Draft, v0.2.0)
            WordMergeCellsTool(self.workspace),
            WordUpdateTableCellTool(self.workspace),
            WordUpdateTableRowColumnTool(self.workspace),
            WordUpdateTableFormatTool(self.workspace),
            # Export tool
            WordExportContentTool(self.workspace),
            # Comment tools
            WordGetCommentsTool(self.workspace),
            WordInsertCommentTool(self.workspace),
            WordDeleteCommentTool(self.workspace),
            WordReplyCommentTool(self.workspace),
            WordResolveCommentTool(self.workspace),
            # Online script escape hatch (OASP word:run:script, issue #87)
            WordRunScriptTool(self.workspace),
        ]

        for tool in word_tools:
            self.tools[tool.name] = tool

        logger.info(f"已注册 {len(word_tools)} 个 Word 工具 | Registered {len(word_tools)} Word tools")

        from office4ai.a2c_smcp.tools.ppt import (
            PptAddSlideTool,
            PptDeleteElementTool,
            PptDeleteSlideTool,
            PptGetChartTool,
            PptGetCurrentSlideElementsTool,
            PptGetSlideElementsTool,
            PptGetSlideInfoTool,
            PptGetSlideLayoutsTool,
            PptGetSlideScreenshotTool,
            PptGotoSlideTool,
            PptInsertChartTool,
            PptInsertImageTool,
            PptInsertShapeTool,
            PptInsertTableTool,
            PptInsertTextTool,
            PptMoveSlideTool,
            PptReorderElementTool,
            PptRunScriptTool,
            PptUpdateChartTool,
            PptUpdateElementTool,
            PptUpdateImageTool,
            PptUpdateTableCellTool,
            PptUpdateTableFormatTool,
            PptUpdateTableRowColumnTool,
            PptUpdateTextBoxTool,
        )

        ppt_tools = [
            # Content retrieval tools
            PptGetCurrentSlideElementsTool(self.workspace),
            PptGetSlideElementsTool(self.workspace),
            PptGetSlideScreenshotTool(self.workspace),
            PptGetSlideInfoTool(self.workspace),
            PptGetSlideLayoutsTool(self.workspace),
            # Content insertion tools
            PptInsertTextTool(self.workspace),
            PptInsertImageTool(self.workspace),
            PptInsertTableTool(self.workspace),
            PptInsertShapeTool(self.workspace),
            # Update operation tools
            PptUpdateTextBoxTool(self.workspace),
            PptUpdateImageTool(self.workspace),
            PptUpdateTableCellTool(self.workspace),
            PptUpdateTableRowColumnTool(self.workspace),
            PptUpdateTableFormatTool(self.workspace),
            PptUpdateElementTool(self.workspace),
            # Chart tools (OASP /ppt Draft, v0.2.0 — Server OOXML)
            PptInsertChartTool(self.workspace),
            PptGetChartTool(self.workspace),
            PptUpdateChartTool(self.workspace),
            # Delete & layout tools
            PptDeleteElementTool(self.workspace),
            PptReorderElementTool(self.workspace),
            # Slide management tools
            PptAddSlideTool(self.workspace),
            PptDeleteSlideTool(self.workspace),
            PptMoveSlideTool(self.workspace),
            PptGotoSlideTool(self.workspace),
            # Online script escape hatch (OASP ppt:run:script, issue #87)
            PptRunScriptTool(self.workspace),
        ]

        for tool in ppt_tools:
            self.tools[tool.name] = tool

        logger.info(f"已注册 {len(ppt_tools)} 个 PPT 工具 | Registered {len(ppt_tools)} PPT tools")

        from office4ai.a2c_smcp.tools.excel import (
            ExcelActivateWorksheetTool,
            ExcelAddConditionalFormatTool,
            ExcelAddTableRowTool,
            ExcelAddWorksheetTool,
            ExcelClearAutoFilterTool,
            ExcelClearConditionalFormatTool,
            ExcelClearRangeTool,
            ExcelCopyRangeTool,
            ExcelDeleteChartTool,
            ExcelDeletePivotTableTool,
            ExcelDeleteRangeTool,
            ExcelDeleteTableRowTool,
            ExcelDeleteWorksheetTool,
            ExcelFindValuesTool,
            ExcelGetChartsTool,
            ExcelGetPivotTablesTool,
            ExcelGetRangeFormatTool,
            ExcelGetRangeTool,
            ExcelGetSelectedRangeTool,
            ExcelGetTablesTool,
            ExcelGetTableTool,
            ExcelGetWorkbookInfoTool,
            ExcelGetWorksheetInfoTool,
            ExcelGetWorksheetsTool,
            ExcelInsertChartTool,
            ExcelInsertPivotTableTool,
            ExcelInsertRangeTool,
            ExcelInsertTableTool,
            ExcelMergeCellsTool,
            ExcelRenameWorksheetTool,
            ExcelRunScriptTool,
            ExcelSetAutoFilterTool,
            ExcelSetFormulaTool,
            ExcelSetRangeFormatTool,
            ExcelSetRangeTool,
            ExcelSortTableTool,
            ExcelUnmergeCellsTool,
            ExcelUpdateChartTool,
        )

        excel_tools = [
            # State-awareness read tools (OASP /excel Draft 0.3.0, issue #18 Foundation)
            ExcelGetWorkbookInfoTool(self.workspace),
            ExcelGetWorksheetInfoTool(self.workspace),
            ExcelGetSelectedRangeTool(self.workspace),
            # Range CRUD + 公式 (OASP /excel Draft 0.3.0, issue #19)
            ExcelGetRangeTool(self.workspace),
            ExcelSetRangeTool(self.workspace),
            ExcelClearRangeTool(self.workspace),
            ExcelCopyRangeTool(self.workspace),
            ExcelDeleteRangeTool(self.workspace),
            ExcelInsertRangeTool(self.workspace),
            ExcelSetFormulaTool(self.workspace),
            # Format / 条件格式 / 合并单元格 (OASP /excel Draft 0.3.0, issue #20)
            ExcelGetRangeFormatTool(self.workspace),
            ExcelSetRangeFormatTool(self.workspace),
            ExcelAddConditionalFormatTool(self.workspace),
            ExcelClearConditionalFormatTool(self.workspace),
            ExcelMergeCellsTool(self.workspace),
            ExcelUnmergeCellsTool(self.workspace),
            # Worksheet 管理 (OASP /excel Draft 0.3.0, issue #21)
            ExcelGetWorksheetsTool(self.workspace),
            ExcelAddWorksheetTool(self.workspace),
            ExcelDeleteWorksheetTool(self.workspace),
            ExcelRenameWorksheetTool(self.workspace),
            ExcelActivateWorksheetTool(self.workspace),
            # Table 操作 (OASP /excel Draft 0.3.0, issue #22)
            ExcelInsertTableTool(self.workspace),
            ExcelGetTableTool(self.workspace),
            ExcelGetTablesTool(self.workspace),
            ExcelAddTableRowTool(self.workspace),
            ExcelDeleteTableRowTool(self.workspace),
            ExcelSortTableTool(self.workspace),
            # Chart 操作 (OASP /excel Draft 0.3.0, issue #23)
            ExcelInsertChartTool(self.workspace),
            ExcelGetChartsTool(self.workspace),
            ExcelUpdateChartTool(self.workspace),
            ExcelDeleteChartTool(self.workspace),
            # PivotTable 操作 (OASP /excel Draft 0.3.0, issue #24)
            ExcelInsertPivotTableTool(self.workspace),
            ExcelGetPivotTablesTool(self.workspace),
            ExcelDeletePivotTableTool(self.workspace),
            # Find & Filter 操作 (OASP /excel Draft 0.3.0, issue #25)
            ExcelFindValuesTool(self.workspace),
            ExcelSetAutoFilterTool(self.workspace),
            ExcelClearAutoFilterTool(self.workspace),
            # Online script escape hatch (OASP excel:run:script, issue #87)
            ExcelRunScriptTool(self.workspace),
        ]

        for tool in excel_tools:
            self.tools[tool.name] = tool

        logger.info(f"已注册 {len(excel_tools)} 个 Excel 工具 | Registered {len(excel_tools)} Excel tools")

        # authoring standalone 工具（milestone #4 · S1）——无需 Add-In 连接，常驻
        from office4ai.a2c_smcp.tools.authoring import OfficeRunScriptTool

        authoring_tools = [
            OfficeRunScriptTool(self.workspace),
        ]
        for tool in authoring_tools:
            self.tools[tool.name] = tool

        logger.info(
            f"已注册 {len(authoring_tools)} 个 authoring 工具 | Registered {len(authoring_tools)} authoring tools",
        )

    def _register_resources(self) -> None:
        """注册资源 | Register resources"""
        from office4ai.a2c_smcp.resources.skill import discover_skill_resources
        from office4ai.a2c_smcp.resources.window import WindowResource

        # 根索引常驻；per-file 子窗口（word/ppt）随 Add-In 连接动态注册（W4b-1 / #64，见
        # _on_doc_connect），取代旧的 per-type 聚合窗口（window://office4ai/word|ppt）。
        root = WindowResource(self.workspace, priority=0, fullscreen=False)
        self.resources[root.base_uri] = root

        # authoring SKILL 能力包（milestone #4 · S3）：扫描 skills_root 下每个含 SKILL.md 的目录，
        # 经 skill:// 的 resources source 模式暴露供 A2C Computer 物化。S3 地基阶段 office/skills/
        # 为空目录 → 无 skill:// 条目；生产 SKILL 由 S4/S5/S6（create/edit/extract）落入该目录。
        for skill in discover_skill_resources(self._skills_root):
            self.resources[skill.base_uri] = skill
            logger.info(f"注册 SKILL 资源 | Registered SKILL resource: {skill.base_uri}")

    # Tool category → Socket.IO namespace（W4a 连接过滤 + per-file 窗口定位用）
    _CATEGORY_NAMESPACE: dict[str, str] = {"word": "/word", "ppt": "/ppt", "excel": "/excel"}
    _ROOT_URI = "window://office4ai"

    async def _async_startup(self) -> None:
        """启动 OfficeWorkspace (Socket.IO Server) | Start OfficeWorkspace"""
        logger.info("启动 OfficeWorkspace | Starting OfficeWorkspace...")
        await self.workspace.start()

        # authoring 运行时依赖探测（milestone #4 · S1）——soffice 缺失仅告警不阻断
        from office4ai.office.authoring import log_soffice_status

        log_soffice_status()

        connection_manager.register_connect_callback(self._on_doc_connect)
        connection_manager.register_disconnect_callback_ns(self._on_doc_disconnect)

        # Wire Server-side OOXML mutations (e.g. /ppt chart engine) to MCP
        # resource_updated notifications so subscribers learn the file changed.
        self.workspace.set_resource_update_callback(self.subscription_manager.notify_fire_and_forget)

        # W4b-2 (#65): flip fullscreen 归属 to the file the AI just operated on.
        self.workspace.set_activity_callback(self._on_doc_activity)

    async def _async_shutdown(self) -> None:
        """停止 OfficeWorkspace (Socket.IO Server) | Stop OfficeWorkspace"""
        logger.info("停止 OfficeWorkspace | Stopping OfficeWorkspace...")
        self.workspace.set_resource_update_callback(None)
        self.workspace.set_activity_callback(None)
        self._activity_order.clear()
        self._activity_seq = 0
        self.subscription_manager.clear()
        await self.workspace.stop()

    # ── 动态工具收敛（W4a / #63）──

    def _is_tool_available(self, tool: BaseTool) -> bool:
        """按 Add-In 连接收敛工具集：常驻工具（requires_connection=False）始终暴露；

        平台工具仅在其对应 namespace 有连接时暴露（无连接 → 移除；仅 Word 连接 → 移除
        PPT/Excel）。契约锚点：spec §「W4 交互契约」契约①。
        """
        if not tool.requires_connection:
            return True
        namespace = self._CATEGORY_NAMESPACE.get(tool.category)
        return bool(namespace and connection_manager.get_clients_by_namespace(namespace))

    def _affected_resource_uris(self, tool: BaseTool, arguments: dict[str, Any]) -> list[str]:
        """成功工具调用 → 通知其操作文件的 per-file 窗口（W4b-1 / W4b-3）。

        用工具入参里的 document_uri（归一化后与注册时同坐标系）定位 per-file 窗口 base_uri
        （word/ppt/excel 均有窗口）；无 document_uri/未知 category（如 authoring）→ 无对应窗口
        （返回空，不通知）。
        """
        doc_uri = arguments.get("document_uri")
        if not isinstance(doc_uri, str) or not doc_uri:
            return []
        namespace = self._CATEGORY_NAMESPACE.get(tool.category)
        if namespace is None:
            return []
        return affected_window_uris(namespace, doc_uri)

    # ── per-file window 动态注册 + 桌面变化通知（W4b-1 / W4a）──

    def _on_doc_connect(self, doc_uri: str, namespace: str) -> None:
        """首个连接：注册该文件的 per-file 窗口，广播桌面变化（工具收敛 + 窗口列表）。"""
        window = create_per_file_window(self.workspace, doc_uri, namespace)
        if window is not None:
            self.resources[window.base_uri] = window
        self._notify_desktop_changed()

    def _on_doc_disconnect(self, doc_uri: str, namespace: str) -> None:
        """完全断连：注销该文件的 per-file 窗口，处理 fullscreen 顺延，广播桌面变化。"""
        base_uri = per_file_window_base_uri(namespace, doc_uri)
        was_fullscreen = False
        if base_uri is not None:
            removed = self.resources.pop(base_uri, None)
            was_fullscreen = bool(isinstance(removed, PerFileWindowResource) and removed.fullscreen)
        self._activity_order.pop(normalize_document_uri(doc_uri), None)

        # 断连收场（契约③）：断连的是 fullscreen 持有者 → 让给剩余已连接文件中最近活跃者；
        # 无最近活跃者则无 fullscreen（根索引主视）。
        if was_fullscreen:
            successor = self._most_recently_active_window()
            self._set_fullscreen_owner(successor.document_uri if successor is not None else None)
        self._notify_desktop_changed()

    def _notify_desktop_changed(self) -> None:
        """连接变化统一通知：W4a tools/list_changed + W4b-1 resources/list_changed + 根索引 resource_updated。"""
        self.subscription_manager.notify_list_changed_fire_and_forget(tools=True, resources=True)
        # 根索引内容变了（列出的 per-file 窗口集变化）→ 通知其订阅者重读。
        self.subscription_manager.notify_fire_and_forget([self._ROOT_URI])

    # ── 单 fullscreen 归属（W4b-2 / #65，关闭 #4）──

    def _on_doc_activity(self, document_uri: str) -> None:
        """AI 最后操作某文件（契约③）：该文件 per-file 窗口置 fullscreen，其余清零。

        竞争仅在已连接（有 window）文件间发生；无对应 window 的文件（office_run_script 产物、
        未连接的盘上编辑）只记台账、不扰动现状 fullscreen。``document_uri`` 已由 workspace 归一化。
        """
        norm = normalize_document_uri(document_uri)
        # 无对应 window → 不参与 fullscreen 竞争，也不记台账（免盘上编辑文件的条目泄漏，
        # 且顺延仅在「有 window 时活跃过」的文件间裁决，不被连接前的陈旧 seq 污染）。
        if self._window_for_document(norm) is None:
            return
        self._activity_seq += 1
        self._activity_order[norm] = self._activity_seq
        changed = self._set_fullscreen_owner(norm)
        if changed:
            # fullscreen 状态编码在 window URI 查询串 → 资源列表表示变化，Computer 重列以重组桌面；
            # 并对翻转的 window 发 resource_updated。工具集未变，不发 tools/list_changed。
            self.subscription_manager.notify_list_changed_fire_and_forget(resources=True)
            self.subscription_manager.notify_fire_and_forget(changed)

    def _per_file_windows(self) -> list[PerFileWindowResource]:
        """当前注册的所有 per-file 窗口资源。"""
        return [r for r in self.resources.values() if isinstance(r, PerFileWindowResource)]

    def _window_for_document(self, norm_uri: str) -> PerFileWindowResource | None:
        """按归一化 document_uri 找到对应 per-file 窗口（无则 None）。"""
        for window in self._per_file_windows():
            if normalize_document_uri(window.document_uri) == norm_uri:
                return window
        return None

    def _set_fullscreen_owner(self, document_uri: str | None) -> list[str]:
        """令 ``document_uri`` 对应窗口成为唯一 fullscreen，其余清零（先清后置 / 原子）。

        ``document_uri=None`` → 全部清零（无 fullscreen）。返回 fullscreen 状态发生翻转的窗口
        base_uri 列表。单趟遍历把每个窗口置为其最终值，遍历结束即满足「至多一个 fullscreen」不变量。
        """
        target = normalize_document_uri(document_uri) if document_uri else None
        changed: list[str] = []
        for window in self._per_file_windows():
            should = target is not None and normalize_document_uri(window.document_uri) == target
            if window.fullscreen != should:
                window.fullscreen = should
                changed.append(window.base_uri)
        return changed

    def _most_recently_active_window(self) -> PerFileWindowResource | None:
        """剩余 per-file 窗口中按台账最近活跃者（都未活跃过则 None）。"""
        best: PerFileWindowResource | None = None
        best_seq = 0  # seq 从 1 起；未活跃过的文件不在台账（get→0），不入选
        for window in self._per_file_windows():
            seq = self._activity_order.get(normalize_document_uri(window.document_uri), 0)
            if seq > best_seq:
                best_seq = seq
                best = window
        return best


# ---------------------------------------------------------------------------
# CLI sub-commands
# ---------------------------------------------------------------------------


def _cmd_setup() -> None:
    """Generate certificates and install CA into system trust store."""
    from office4ai.certs.generator import generate_ca, generate_server_cert

    cert_dir = get_cert_dir()
    status = validate_certs(cert_dir)

    if status == CertStatus.ALL_VALID:
        info = get_cert_expiry_info(cert_dir)
        print(f"Certificates are already valid at {cert_dir}")
        print(f"  CA valid until:     {info.get('ca_valid_until', 'N/A')}")
        print(f"  Server valid until: {info.get('server_valid_until', 'N/A')}")
        return

    # Determine what needs to happen
    need_ca = status in (CertStatus.NO_CERTS, CertStatus.CA_EXPIRED, CertStatus.INVALID)
    need_server = True  # always regenerate server cert when setup runs

    print()
    print("Office4AI Certificate Setup")
    print("=" * 40)
    if need_ca:
        print("This will:")
        print("  1. Generate a local CA certificate (Name Constraints: localhost/127.0.0.1 only)")
        print("  2. Generate a server certificate for localhost and 127.0.0.1")
        print("  3. Install the CA into your system trust store (requires admin privileges)")
    else:
        print("CA certificate is still valid. This will:")
        print("  1. Regenerate the server certificate (no admin privileges needed)")
    print(f"\nCertificate location: {cert_dir}")
    print()

    answer = input("Proceed? [y/N]: ").strip().lower()
    if answer != "y":
        print("Aborted.")
        return

    if need_ca:
        ca_cert, ca_key = generate_ca(cert_dir)
        print("CA certificate generated (valid 10 years)")
    else:
        # Load existing CA
        from cryptography import x509
        from cryptography.hazmat.primitives.serialization import load_pem_private_key

        from office4ai.certs.paths import CA_CERT_FILE, CA_KEY_FILE

        ca_cert = x509.load_pem_x509_certificate((cert_dir / CA_CERT_FILE).read_bytes())
        ca_key = load_pem_private_key((cert_dir / CA_KEY_FILE).read_bytes(), password=None)  # type: ignore[assignment]

    if need_server:
        generate_server_cert(cert_dir, ca_cert, ca_key)
        print("Server certificate generated (valid 825 days)")

    # Install CA to system trust store (only if CA was regenerated)
    if need_ca:
        from office4ai.certs.paths import CA_CERT_FILE

        ca_cert_path = cert_dir / CA_CERT_FILE
        trust_store = get_trust_store()

        print("Installing CA to system trust store...")
        if trust_store.install(ca_cert_path):
            print("CA installed to system trust store")
        else:
            print("Failed to install CA to system trust store.")
            print("You can install it manually:")
            print(f"  {trust_store.get_manual_install_command(ca_cert_path)}")

    print()
    print("Setup complete! Start the server with: office4ai-mcp serve")


def _cmd_cleanup() -> None:
    """Remove CA from system trust store and delete certificate files."""
    cert_dir = get_cert_dir()

    if not cert_dir.exists():
        print(f"No certificates found at {cert_dir}")
        return

    print()
    print("Office4AI Certificate Cleanup")
    print("=" * 40)
    print("This will:")
    print("  1. Remove CA from system trust store (requires admin privileges)")
    print(f"  2. Delete all certificate files from {cert_dir}")
    print()

    answer = input("Proceed? [y/N]: ").strip().lower()
    if answer != "y":
        print("Aborted.")
        return

    from office4ai.certs.paths import CA_CERT_FILE

    ca_cert_path = cert_dir / CA_CERT_FILE
    trust_store = get_trust_store()

    if ca_cert_path.exists():
        if trust_store.uninstall(ca_cert_path):
            print("CA removed from system trust store")
        else:
            print("Failed to remove CA from system trust store.")
            print("You can remove it manually:")
            print(f"  {trust_store.get_manual_uninstall_command(ca_cert_path)}")

    shutil.rmtree(cert_dir)
    print("Certificate files deleted")
    print()
    print("Cleanup complete")


async def async_main() -> None:
    cert_dir = get_cert_dir()
    status = validate_certs(cert_dir)

    if status != CertStatus.ALL_VALID:
        if status == CertStatus.NO_CERTS:
            logger.error(f"SSL certificates not found at {cert_dir}")
        elif status == CertStatus.SERVER_EXPIRED:
            info = get_cert_expiry_info(cert_dir)
            logger.error(f"Server certificate expired ({info.get('server_valid_until', 'unknown')})")
        elif status == CertStatus.CA_EXPIRED:
            info = get_cert_expiry_info(cert_dir)
            logger.error(f"CA certificate expired ({info.get('ca_valid_until', 'unknown')})")
        else:
            logger.error(f"Certificate validation failed: {status.value}")

        logger.error("Run `office4ai-mcp setup` to generate and install certificates.")
        sys.exit(1)

    info = get_cert_expiry_info(cert_dir)
    logger.info(f"Loaded certificates from {cert_dir}")
    logger.info(f"  CA valid until:     {info.get('ca_valid_until', 'N/A')}")
    logger.info(f"  Server valid until: {info.get('server_valid_until', 'N/A')}")

    config = MCPServerConfig()
    logger.info(
        f"启动 MCP Server | Starting MCP Server: transport={config.transport}, host={config.host}, port={config.port}",
    )

    server = OfficeMCPServer(config, cert_dir=cert_dir)
    await server.run()


def main() -> None:
    from office4ai.logging import setup_logging

    setup_logging()

    # Extract subcommand before confz sees sys.argv
    # Use simple argv inspection: first non-flag arg is the subcommand
    args = sys.argv[1:]
    subcommand = "serve"

    if args and args[0] in ("serve", "setup", "cleanup"):
        subcommand = args[0]
        sys.argv = [sys.argv[0]] + args[1:]

    if subcommand == "setup":
        _cmd_setup()
    elif subcommand == "cleanup":
        _cmd_cleanup()
    else:
        asyncio.run(async_main())


if __name__ == "__main__":
    main()
