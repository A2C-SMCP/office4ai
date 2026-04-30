# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

import asyncio
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

    def __init__(self, config: MCPServerConfig, cert_dir: Path | None = None) -> None:
        # 同步：创建 workspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
            cert_dir=cert_dir,
        )
        super().__init__(config=config, server_name="office4ai")

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

        # Excel 工具 (未来) | Excel tools (future)

    def _register_resources(self) -> None:
        """注册资源 | Register resources"""
        from office4ai.a2c_smcp.resources.ppt_window import PptWindowResource
        from office4ai.a2c_smcp.resources.window import WindowResource
        from office4ai.a2c_smcp.resources.word_window import WordWindowResource

        root = WindowResource(self.workspace, priority=0, fullscreen=False)
        word = WordWindowResource(self.workspace, priority=50, fullscreen=False)
        ppt = PptWindowResource(self.workspace, priority=50, fullscreen=False)

        self.resources[root.base_uri] = root
        self.resources[word.base_uri] = word
        self.resources[ppt.base_uri] = ppt

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
