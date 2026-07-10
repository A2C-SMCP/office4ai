"""MCP Resource Subscription Manager.

Tracks which ServerSessions have subscribed to which resource URIs,
and sends ``resource_updated`` notifications when content changes.

W4a/W4b-1 (#63/#64) additionally broadcast ``tools/list_changed`` /
``resources/list_changed`` to all active client sessions when Add-In connections
change (dynamic tool convergence + per-file windows). List-changed是全局通知（不绑定
某个订阅 URI），故单独维护一份「活跃会话集」并广播。
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable, Coroutine
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from mcp.server.session import ServerSession

logger = logging.getLogger(__name__)


class SubscriptionManager:
    """Manages resource subscriptions and dispatches update notifications."""

    def __init__(self) -> None:
        self._subscriptions: dict[str, set[ServerSession]] = {}
        # 活跃客户端会话集（订阅或调用过 list_* 的会话），用于广播 list_changed 全局通知。
        self._sessions: set[ServerSession] = set()
        # fire-and-forget 通知任务的强引用（事件循环只持弱引用，无强引用可能被 GC 中途回收）。
        self._pending: set[asyncio.Task[None]] = set()

    def _spawn(self, loop: asyncio.AbstractEventLoop, coro: Coroutine[Any, Any, None]) -> None:
        """Schedule *coro* on *loop*, holding a strong ref until it completes."""
        task = loop.create_task(coro)
        self._pending.add(task)
        task.add_done_callback(self._pending.discard)

    def track_session(self, session: ServerSession) -> None:
        """Record *session* as an active client (for list_changed broadcasts)."""
        self._sessions.add(session)

    def subscribe(self, uri: str, session: ServerSession) -> None:
        """Register *session* as a subscriber of *uri*."""
        self._subscriptions.setdefault(uri, set()).add(session)
        self._sessions.add(session)
        logger.debug("Session subscribed to %s (total: %d)", uri, len(self._subscriptions[uri]))

    def unsubscribe(self, uri: str, session: ServerSession) -> None:
        """Remove *session* from subscribers of *uri*."""
        sessions = self._subscriptions.get(uri)
        if sessions:
            sessions.discard(session)
            if not sessions:
                del self._subscriptions[uri]
        logger.debug("Session unsubscribed from %s", uri)

    async def notify(self, uri: str) -> None:
        """Send ``resource_updated`` to all sessions subscribed to *uri*.

        Dead sessions (broken pipe, closed connection) are automatically
        removed on failure (lazy cleanup).
        """
        from pydantic import AnyUrl

        sessions = self._subscriptions.get(uri)
        if not sessions:
            return

        dead: list[ServerSession] = []
        any_url = AnyUrl(uri)

        for session in sessions:
            try:
                await session.send_resource_updated(any_url)
            except Exception:
                logger.debug("Failed to notify session for %s, removing dead session", uri)
                dead.append(session)

        for s in dead:
            sessions.discard(s)
        if not sessions:
            del self._subscriptions[uri]

    async def notify_many(self, uris: list[str]) -> None:
        """Send ``resource_updated`` for each URI in *uris*."""
        for uri in uris:
            await self.notify(uri)

    def notify_fire_and_forget(self, uris: list[str]) -> None:
        """Schedule :meth:`notify_many` from a synchronous context.

        Uses the running event loop's ``create_task`` to bridge sync → async.
        """
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            logger.warning("No running event loop; skipping subscription notification")
            return
        self._spawn(loop, self.notify_many(uris))

    # ── list_changed 广播（W4a/W4b-1 · #63/#64）| list_changed broadcast ──

    def _forget_session(self, session: ServerSession) -> None:
        """Drop a dead session from the active set and all subscription buckets."""
        self._sessions.discard(session)
        for subs in self._subscriptions.values():
            subs.discard(session)

    async def _broadcast(self, send: Callable[[ServerSession], Awaitable[None]]) -> None:
        """Send a notification to every active session; drop dead ones (lazy cleanup)."""
        dead: list[ServerSession] = []
        for session in list(self._sessions):
            try:
                await send(session)
            except Exception:
                logger.debug("Failed to broadcast to session, removing dead session")
                dead.append(session)
        for s in dead:
            self._forget_session(s)

    async def notify_tool_list_changed(self) -> None:
        """Broadcast ``notifications/tools/list_changed`` to all active sessions."""
        await self._broadcast(lambda s: s.send_tool_list_changed())

    async def notify_resource_list_changed(self) -> None:
        """Broadcast ``notifications/resources/list_changed`` to all active sessions."""
        await self._broadcast(lambda s: s.send_resource_list_changed())

    def notify_list_changed_fire_and_forget(self, *, tools: bool = False, resources: bool = False) -> None:
        """Schedule tool/resource list_changed broadcasts from a synchronous context."""
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            logger.warning("No running event loop; skipping list_changed broadcast")
            return
        if tools:
            self._spawn(loop, self.notify_tool_list_changed())
        if resources:
            self._spawn(loop, self.notify_resource_list_changed())

    def clear(self) -> None:
        """Remove all subscriptions and tracked sessions."""
        self._subscriptions.clear()
        self._sessions.clear()
        logger.debug("All subscriptions cleared")
