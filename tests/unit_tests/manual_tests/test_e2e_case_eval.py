"""
manual_tests/excel/e2e_case.py 断言纯函数单测（issue #82 / oasp#17）

覆盖 parse_error_details（与 office_workspace.format_wire_error 渲染格式配对的
反解析）与 evaluate_error_case（错误码 + details 子集严格判定）。xfail 三态属
_run_single 的编排逻辑，其判定核心即 evaluate_error_case，在此一并覆盖边界。
"""

from manual_tests.excel.e2e_case import evaluate_error_case, parse_error_details
from office4ai.environment.workspace.office_workspace import format_wire_error


class TestParseErrorDetails:
    """parse_error_details() — 从摊平错误字符串反解析 details"""

    def test_roundtrip_with_format_wire_error(self) -> None:
        """与生产侧渲染函数往返一致（单一渲染格式的配对保证）"""
        wire = {"code": "3010", "message": "Not found", "details": {"kind": "worksheet", "name": "Ghost"}}
        err = format_wire_error(wire)
        assert parse_error_details(err) == {"kind": "worksheet", "name": "Ghost"}

    def test_no_details_suffix(self) -> None:
        assert parse_error_details("3009: Invalid range") is None

    def test_malformed_json(self) -> None:
        assert parse_error_details("3010: x (details: {broken)") is None

    def test_chinese_value(self) -> None:
        err = format_wire_error({"code": "3010", "message": "x", "details": {"kind": "worksheet", "name": "工作表1"}})
        parsed = parse_error_details(err)
        assert parsed is not None
        assert parsed["name"] == "工作表1"


class TestEvaluateErrorCase:
    """evaluate_error_case() — 错误码路径严格判定"""

    ERR_3010 = '3010: Worksheet not found (details: {"kind": "worksheet", "name": "Ghost"})'

    def test_code_hit(self) -> None:
        assert evaluate_error_case(False, "3009: Invalid range", "3009") is True

    def test_code_miss(self) -> None:
        """Add-In 未接线仍返 3000 → 严格判定不过（由 xfail_reason 兜底为 XFAIL）"""
        assert evaluate_error_case(False, "3000: Document error", "3010") is False

    def test_success_never_passes(self) -> None:
        """本应失败却成功 → 必须 False（xfail 也不得掩盖此回归）"""
        assert evaluate_error_case(True, None, "3010") is False

    def test_none_error(self) -> None:
        assert evaluate_error_case(False, None, "3010") is False

    def test_details_subset_match(self) -> None:
        assert evaluate_error_case(False, self.ERR_3010, "3010", {"kind": "worksheet"}) is True

    def test_details_full_match(self) -> None:
        assert evaluate_error_case(False, self.ERR_3010, "3010", {"kind": "worksheet", "name": "Ghost"}) is True

    def test_details_value_mismatch(self) -> None:
        assert evaluate_error_case(False, self.ERR_3010, "3010", {"kind": "table"}) is False

    def test_details_expected_but_absent(self) -> None:
        """期望 details 但错误串无 details 后缀 → 不过"""
        assert evaluate_error_case(False, "3010: Worksheet not found", "3010", {"kind": "worksheet"}) is False

    def test_details_not_expected_ignores_actual(self) -> None:
        """未声明 expect_details 时不校验实收 details（只看码）"""
        assert evaluate_error_case(False, self.ERR_3010, "3010") is True
