"""
OASP Protocol Versioning

协议版本号语义与兼容性判定（OASP 0.3.0 引入的连接期版本握手地基）。

参考规范：
- conventions.md#versioning / conventions.md#compatibility-rule
- connection.md#protocol-version-handshake

版本号语义为 SemVer ``MAJOR.MINOR.PATCH``，单一事实源为 ``office4ai.__version__``
（同步自 ``pyproject.toml`` 的 ``version`` 字段，由 bump-my-version 管理）。
"""

from __future__ import annotations

from dataclasses import dataclass

from office4ai import __version__


@dataclass(frozen=True)
class OaspVersion:
    """
    OASP 协议版本号（不可变值对象）。

    Example:
        >>> OaspVersion.parse("0.3.0")
        OaspVersion(major=0, minor=3, patch=0)
        >>> str(OaspVersion(0, 3, 0))
        '0.3.0'
    """

    major: int
    minor: int
    patch: int

    @classmethod
    def parse(cls, s: str) -> OaspVersion:
        """
        从 ``MAJOR.MINOR.PATCH`` 字符串解析版本号。

        Args:
            s: SemVer 版本字符串（如 ``"0.3.0"``）

        Returns:
            解析后的 ``OaspVersion``

        Raises:
            ValueError: 格式非法（段数不为 3 或非整数段）
        """
        parts = s.split(".")
        if len(parts) != 3:
            raise ValueError(f"Invalid version: {s}")
        try:
            major, minor, patch = (int(p) for p in parts)
        except ValueError as exc:
            raise ValueError(f"Invalid version: {s}") from exc
        return cls(major, minor, patch)

    def __str__(self) -> str:
        return f"{self.major}.{self.minor}.{self.patch}"


def is_compatible(client: OaspVersion, server: OaspVersion) -> bool:
    """
    判定握手中声明 ``oaspVersion`` 的 AddIn（client）是否与 Server 兼容。

    - **v0.x（MAJOR=0，不稳定阶段）**：任何 MINOR 都可能是破坏性变更，
      故 ``MAJOR.MINOR`` 必须严格相等（PATCH 可自由差异）。
    - **v1.0+（MAJOR≥1，稳定阶段）**：MAJOR 必须相等，且 Server MINOR ≥ Client MINOR
      （较新的 Server 向后兼容较旧的 AddIn）。

    判定公式与 A2C-SMCP 协议一致，升级次序相同：**Server 先于 AddIn 升级**。

    规范：conventions.md#compatibility-rule
    """
    if client.major != server.major:
        return False
    if client.major == 0:
        return client.minor == server.minor  # v0.x 严格匹配 MINOR
    return client.minor <= server.minor  # v1.0+ Server 向后兼容较旧 AddIn


# ---------------------------------------------------------------------------
# Server 版本常量（单一事实源：office4ai.__version__）
# ---------------------------------------------------------------------------

#: 本 Server 实现的 OASP 协议版本（用于兼容性判定与 connection:established 诊断字段）
SERVER_VERSION: OaspVersion = OaspVersion.parse(__version__)

#: Server 支持的最低 / 最高版本，仅用于不兼容时回送给 AddIn 的诊断字段。
#: v0.x 阶段实际放行与否由 is_compatible 的严格 MAJOR.MINOR 判定，min/max 不参与决策。
SERVER_MIN_SUPPORTED: OaspVersion = OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor, 0)
SERVER_MAX_SUPPORTED: OaspVersion = OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor, 999)
