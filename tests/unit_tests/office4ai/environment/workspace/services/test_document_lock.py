"""Unit tests for the per-document async lock manager."""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.services.document_lock import DocumentLockManager


@pytest.mark.asyncio
async def test_acquire_serializes_concurrent_writes() -> None:
    """Two coroutines on the same document_uri must run sequentially under the lock."""
    mgr = DocumentLockManager()
    uri = "file:///tmp/foo.pptx"
    order: list[str] = []

    async def writer(label: str, hold: float) -> None:
        async with mgr.acquire(uri):
            order.append(f"{label}:enter")
            await asyncio.sleep(hold)
            order.append(f"{label}:exit")

    await asyncio.gather(writer("A", 0.05), writer("B", 0.0))
    # Whoever entered first must exit before the other enters.
    assert order[0].endswith(":enter")
    assert order[1].endswith(":exit")
    assert order[2].endswith(":enter")
    assert order[3].endswith(":exit")
    assert order[0].split(":")[0] == order[1].split(":")[0]
    assert order[2].split(":")[0] == order[3].split(":")[0]


@pytest.mark.asyncio
async def test_different_documents_run_in_parallel() -> None:
    """Locks for distinct URIs must be independent."""
    mgr = DocumentLockManager()
    in_critical: list[str] = []

    async def writer(uri: str, label: str) -> None:
        async with mgr.acquire(uri):
            in_critical.append(label)
            await asyncio.sleep(0.05)
            in_critical.append(f"{label}:done")

    await asyncio.gather(
        writer("file:///tmp/a.pptx", "A"),
        writer("file:///tmp/b.pptx", "B"),
    )
    # Both writers should overlap: both labels appear before either *:done.
    enter_indices = [i for i, x in enumerate(in_critical) if not x.endswith(":done")]
    done_indices = [i for i, x in enumerate(in_critical) if x.endswith(":done")]
    assert max(enter_indices) < min(done_indices)


@pytest.mark.asyncio
async def test_normalized_uris_share_lock() -> None:
    """Different surface forms of the same path must collide on a single lock."""
    mgr = DocumentLockManager()
    order: list[str] = []

    async def writer(label: str, uri: str, hold: float) -> None:
        async with mgr.acquire(uri):
            order.append(f"{label}:enter")
            await asyncio.sleep(hold)
            order.append(f"{label}:exit")

    await asyncio.gather(
        writer("A", "file:///Users/x/foo.pptx", 0.05),
        writer("B", "file:///Users/x/foo.pptx", 0.0),
    )
    assert order[0].endswith(":enter")
    assert order[1].endswith(":exit")
    assert order[2].endswith(":enter")
    assert order[3].endswith(":exit")


@pytest.mark.asyncio
async def test_is_locked_reports_state() -> None:
    mgr = DocumentLockManager()
    uri = "file:///tmp/c.pptx"
    assert not mgr.is_locked(uri)
    async with mgr.acquire(uri):
        assert mgr.is_locked(uri)
    assert not mgr.is_locked(uri)
