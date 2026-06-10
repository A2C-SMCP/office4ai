"""
Test OASP protocol versioning

测试 OaspVersion 解析与 is_compatible 兼容性判定（OASP 0.3.0 协议版本握手地基）。
规范：conventions.md#versioning / conventions.md#compatibility-rule。
"""

import pytest

from office4ai import __version__
from office4ai.environment.workspace.socketio.versioning import (
    SERVER_MAX_SUPPORTED,
    SERVER_MIN_SUPPORTED,
    SERVER_VERSION,
    OaspVersion,
    is_compatible,
)


class TestOaspVersionParse:
    """OaspVersion.parse / __str__"""

    def test_parse_valid(self) -> None:
        v = OaspVersion.parse("0.3.0")
        assert (v.major, v.minor, v.patch) == (0, 3, 0)

    def test_parse_multi_digit(self) -> None:
        v = OaspVersion.parse("12.34.567")
        assert (v.major, v.minor, v.patch) == (12, 34, 567)

    def test_str_roundtrip(self) -> None:
        assert str(OaspVersion.parse("1.2.3")) == "1.2.3"

    def test_frozen_hashable(self) -> None:
        # frozen dataclass：可作为 dict key / set 元素
        assert OaspVersion(0, 3, 0) == OaspVersion(0, 3, 0)
        assert len({OaspVersion(0, 3, 0), OaspVersion(0, 3, 0)}) == 1

    @pytest.mark.parametrize("bad", ["0.3", "0.3.0.1", "", "abc", "0.x.0", "1..2", "0.3.0-rc1"])
    def test_parse_invalid_raises_valueerror(self, bad: str) -> None:
        with pytest.raises(ValueError):
            OaspVersion.parse(bad)


class TestIsCompatible:
    """is_compatible —— v0.x 严格 MAJOR.MINOR，v1.0+ Server 向后兼容"""

    def test_v0x_same_minor_compatible(self) -> None:
        assert is_compatible(OaspVersion(0, 3, 0), OaspVersion(0, 3, 0)) is True

    def test_v0x_same_minor_diff_patch_compatible(self) -> None:
        # v0.x：PATCH 可自由差异
        assert is_compatible(OaspVersion(0, 3, 7), OaspVersion(0, 3, 0)) is True
        assert is_compatible(OaspVersion(0, 3, 0), OaspVersion(0, 3, 9)) is True

    def test_v0x_diff_minor_incompatible(self) -> None:
        # v0.x：0.2 与 0.3 互拒
        assert is_compatible(OaspVersion(0, 2, 0), OaspVersion(0, 3, 0)) is False
        assert is_compatible(OaspVersion(0, 4, 0), OaspVersion(0, 3, 0)) is False

    def test_diff_major_incompatible(self) -> None:
        assert is_compatible(OaspVersion(1, 0, 0), OaspVersion(0, 3, 0)) is False
        assert is_compatible(OaspVersion(0, 3, 0), OaspVersion(1, 3, 0)) is False

    def test_v1plus_server_minor_ge_client(self) -> None:
        # v1.0+：Server MINOR ≥ Client MINOR → 兼容（较新 Server 向后兼容较旧 AddIn）
        assert is_compatible(OaspVersion(1, 2, 0), OaspVersion(1, 5, 0)) is True
        assert is_compatible(OaspVersion(1, 5, 0), OaspVersion(1, 5, 0)) is True

    def test_v1plus_client_minor_gt_server_incompatible(self) -> None:
        # v1.0+：Client 比 Server 新 → 不兼容
        assert is_compatible(OaspVersion(1, 6, 0), OaspVersion(1, 5, 0)) is False


class TestServerConstants:
    """Server 版本常量 —— 单一事实源 office4ai.__version__"""

    def test_server_version_matches_package(self) -> None:
        assert str(SERVER_VERSION) == __version__

    def test_server_version_is_030(self) -> None:
        assert (SERVER_VERSION.major, SERVER_VERSION.minor) == (0, 3)

    def test_min_max_window(self) -> None:
        assert SERVER_MIN_SUPPORTED == OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor, 0)
        assert SERVER_MAX_SUPPORTED == OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor, 999)

    def test_server_self_compatible(self) -> None:
        assert is_compatible(SERVER_VERSION, SERVER_VERSION) is True
