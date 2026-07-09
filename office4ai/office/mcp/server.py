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

from loguru import logger

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.certs.paths import get_cert_dir
from office4ai.certs.trust_store import get_trust_store
from office4ai.certs.validator import CertStatus, get_cert_expiry_info, validate_certs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


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
        from office4ai.a2c_smcp.resources.ppt_window import PptWindowResource
        from office4ai.a2c_smcp.resources.skill import discover_skill_resources
        from office4ai.a2c_smcp.resources.window import WindowResource
        from office4ai.a2c_smcp.resources.word_window import WordWindowResource

        root = WindowResource(self.workspace, priority=0, fullscreen=False)
        word = WordWindowResource(self.workspace, priority=50, fullscreen=False)
        ppt = PptWindowResource(self.workspace, priority=50, fullscreen=False)

        self.resources[root.base_uri] = root
        self.resources[word.base_uri] = word
        self.resources[ppt.base_uri] = ppt

        # authoring SKILL 能力包（milestone #4 · S3）：扫描 skills_root 下每个含 SKILL.md 的目录，
        # 经 skill:// 的 resources source 模式暴露供 A2C Computer 物化。S3 地基阶段 office/skills/
        # 为空目录 → 无 skill:// 条目；生产 SKILL 由 S4/S5/S6（create/edit/extract）落入该目录。
        for skill in discover_skill_resources(self._skills_root):
            self.resources[skill.base_uri] = skill
            logger.info(f"注册 SKILL 资源 | Registered SKILL resource: {skill.base_uri}")

    # Namespace → resource URIs mapping
    _NAMESPACE_URI_MAP: dict[str, str] = {
        "/word": "window://office4ai/word",
        "/ppt": "window://office4ai/ppt",
        "/excel": "window://office4ai/excel",
    }
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

    async def _async_shutdown(self) -> None:
        """停止 OfficeWorkspace (Socket.IO Server) | Stop OfficeWorkspace"""
        logger.info("停止 OfficeWorkspace | Stopping OfficeWorkspace...")
        self.workspace.set_resource_update_callback(None)
        self.subscription_manager.clear()
        await self.workspace.stop()

    def _on_doc_connect(self, doc_uri: str, namespace: str) -> None:
        """Bridge document connect event to MCP resource subscription notifications."""
        uris = self._namespace_to_uris(namespace)
        self.subscription_manager.notify_fire_and_forget(uris)

    def _on_doc_disconnect(self, doc_uri: str, namespace: str) -> None:
        """Bridge document disconnect event to MCP resource subscription notifications."""
        uris = self._namespace_to_uris(namespace)
        self.subscription_manager.notify_fire_and_forget(uris)

    def _namespace_to_uris(self, namespace: str) -> list[str]:
        """Map a Socket.IO namespace to affected resource URIs (platform + root)."""
        uris = [self._ROOT_URI]
        platform_uri = self._NAMESPACE_URI_MAP.get(namespace)
        if platform_uri:
            uris.append(platform_uri)
        return uris


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
