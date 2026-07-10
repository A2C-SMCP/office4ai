"""
Document-level async lock manager.

Used by the OOXML chart engine to serialize concurrent writes against the same
document. The OASP /ppt chart events drive a base64 slide round-trip (export →
mutate in-memory → re-insert); two concurrent writes would corrupt the package.

Usage::

    async with document_lock_manager.acquire(document_uri):
        ...  # exclusive OOXML write

The lock is keyed by the *normalized* document URI so that
``file:///path/foo.pptx`` and ``file://localhost/path/foo.pptx`` share the same
lock.
"""

from __future__ import annotations

import asyncio
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from office4ai.environment.workspace.socketio.services.connection_manager import normalize_document_uri


class DocumentLockManager:
    """Async lock registry keyed by normalized document URI."""

    def __init__(self) -> None:
        self._locks: dict[str, asyncio.Lock] = {}
        self._registry_lock = asyncio.Lock()

    async def _get_or_create(self, document_uri: str) -> asyncio.Lock:
        key = normalize_document_uri(document_uri)
        async with self._registry_lock:
            lock = self._locks.get(key)
            if lock is None:
                lock = asyncio.Lock()
                self._locks[key] = lock
            return lock

    @asynccontextmanager
    async def acquire(self, document_uri: str) -> AsyncIterator[None]:
        """Acquire the per-document lock for the duration of the ``async with``."""
        lock = await self._get_or_create(document_uri)
        async with lock:
            yield

    def is_locked(self, document_uri: str) -> bool:
        """Best-effort check (mainly for diagnostics / tests)."""
        key = normalize_document_uri(document_uri)
        lock = self._locks.get(key)
        return bool(lock and lock.locked())


# Module-level singleton — chart engine and tests share this instance.
document_lock_manager = DocumentLockManager()
