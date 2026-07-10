"""Tests for SubscriptionManager."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.subscriptions import SubscriptionManager


def _make_session() -> MagicMock:
    """Create a mock ServerSession with resource_updated + list_changed senders."""
    session = MagicMock()
    session.send_resource_updated = AsyncMock()
    session.send_tool_list_changed = AsyncMock()
    session.send_resource_list_changed = AsyncMock()
    return session


class TestSubscriptionManager:
    def setup_method(self) -> None:
        self.mgr = SubscriptionManager()

    def test_subscribe_and_unsubscribe(self) -> None:
        session = _make_session()
        self.mgr.subscribe("window://office4ai/word", session)
        assert "window://office4ai/word" in self.mgr._subscriptions
        assert session in self.mgr._subscriptions["window://office4ai/word"]

        self.mgr.unsubscribe("window://office4ai/word", session)
        assert "window://office4ai/word" not in self.mgr._subscriptions

    def test_unsubscribe_nonexistent(self) -> None:
        """Unsubscribing a URI that was never subscribed should not raise."""
        session = _make_session()
        self.mgr.unsubscribe("window://office4ai/word", session)

    @pytest.mark.asyncio
    async def test_notify_sends_to_all_subscribers(self) -> None:
        s1 = _make_session()
        s2 = _make_session()
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, s1)
        self.mgr.subscribe(uri, s2)

        await self.mgr.notify(uri)

        s1.send_resource_updated.assert_awaited_once()
        s2.send_resource_updated.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_notify_no_subscribers(self) -> None:
        """notify on an unsubscribed URI should be a no-op."""
        await self.mgr.notify("window://office4ai/word")

    @pytest.mark.asyncio
    async def test_notify_removes_dead_session(self) -> None:
        alive = _make_session()
        dead = _make_session()
        dead.send_resource_updated.side_effect = Exception("connection lost")
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, alive)
        self.mgr.subscribe(uri, dead)

        await self.mgr.notify(uri)

        # Dead session should have been removed
        assert dead not in self.mgr._subscriptions.get(uri, set())
        # Alive session remains
        assert alive in self.mgr._subscriptions[uri]

    @pytest.mark.asyncio
    async def test_notify_removes_uri_when_all_sessions_dead(self) -> None:
        dead = _make_session()
        dead.send_resource_updated.side_effect = Exception("closed")
        uri = "window://office4ai/ppt"

        self.mgr.subscribe(uri, dead)
        await self.mgr.notify(uri)

        assert uri not in self.mgr._subscriptions

    @pytest.mark.asyncio
    async def test_notify_many(self) -> None:
        s = _make_session()
        self.mgr.subscribe("window://office4ai/word", s)
        self.mgr.subscribe("window://office4ai", s)

        await self.mgr.notify_many(["window://office4ai/word", "window://office4ai"])

        assert s.send_resource_updated.await_count == 2

    def test_clear(self) -> None:
        s = _make_session()
        self.mgr.subscribe("window://office4ai/word", s)
        self.mgr.subscribe("window://office4ai/ppt", s)

        self.mgr.clear()
        assert len(self.mgr._subscriptions) == 0

    def test_multiple_sessions_same_uri(self) -> None:
        s1 = _make_session()
        s2 = _make_session()
        uri = "window://office4ai/word"

        self.mgr.subscribe(uri, s1)
        self.mgr.subscribe(uri, s2)

        # Unsubscribe one; the other should remain
        self.mgr.unsubscribe(uri, s1)
        assert uri in self.mgr._subscriptions
        assert s2 in self.mgr._subscriptions[uri]


class TestListChangedBroadcast:
    """W4a/W4b-1（#63/#64）：list_changed 全局广播 + 会话跟踪。"""

    def setup_method(self) -> None:
        self.mgr = SubscriptionManager()

    def test_track_session_and_subscribe_populate_sessions(self) -> None:
        s1 = _make_session()
        s2 = _make_session()
        self.mgr.track_session(s1)
        self.mgr.subscribe("window://office4ai", s2)  # subscribe 也跟踪
        assert self.mgr._sessions == {s1, s2}

    @pytest.mark.asyncio
    async def test_notify_tool_list_changed_broadcasts_to_all(self) -> None:
        s1, s2 = _make_session(), _make_session()
        self.mgr.track_session(s1)
        self.mgr.track_session(s2)
        await self.mgr.notify_tool_list_changed()
        s1.send_tool_list_changed.assert_awaited_once()
        s2.send_tool_list_changed.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_notify_resource_list_changed_broadcasts_to_all(self) -> None:
        s = _make_session()
        self.mgr.track_session(s)
        await self.mgr.notify_resource_list_changed()
        s.send_resource_list_changed.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_broadcast_removes_dead_session(self) -> None:
        alive, dead = _make_session(), _make_session()
        dead.send_tool_list_changed.side_effect = Exception("closed")
        self.mgr.subscribe("window://office4ai", alive)
        self.mgr.subscribe("window://office4ai", dead)
        self.mgr.track_session(alive)
        self.mgr.track_session(dead)

        await self.mgr.notify_tool_list_changed()

        # 死会话被移出会话集与订阅桶
        assert dead not in self.mgr._sessions
        assert alive in self.mgr._sessions
        assert dead not in self.mgr._subscriptions.get("window://office4ai", set())

    @pytest.mark.asyncio
    async def test_fire_and_forget_schedules_both(self) -> None:
        s = _make_session()
        self.mgr.track_session(s)
        self.mgr.notify_list_changed_fire_and_forget(tools=True, resources=True)
        # 让调度的 task 执行
        import asyncio

        await asyncio.sleep(0)
        await asyncio.sleep(0)
        s.send_tool_list_changed.assert_awaited_once()
        s.send_resource_list_changed.assert_awaited_once()

    def test_fire_and_forget_no_loop_is_noop(self) -> None:
        # 无运行事件循环时安全 no-op（不抛）
        s = _make_session()
        self.mgr.track_session(s)
        self.mgr.notify_list_changed_fire_and_forget(tools=True)

    def test_clear_drops_sessions(self) -> None:
        s = _make_session()
        self.mgr.subscribe("window://office4ai", s)
        self.mgr.track_session(s)
        self.mgr.clear()
        assert self.mgr._sessions == set()
        assert self.mgr._subscriptions == {}
