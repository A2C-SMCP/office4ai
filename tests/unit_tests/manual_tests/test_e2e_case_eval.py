"""
错误码用例判定纯函数单测（issue #82 / oasp#17；#90 判定上提为三命名空间共享后扩充）

覆盖三层：
- ``parse_error_details`` —— 与 ``office_workspace.format_wire_error`` 渲染格式配对的反解析
- ``evaluate_error_case`` —— 错误码 + details 子集的**严格**判定
- ``judge_error_case``   —— xfail 三态编排（#90 从 _run_single 内联块抽出后可直接测）

导入路径刻意混用：``judge_error_case`` 取自新的共享模块 ``manual_tests.error_case``，而
``evaluate_error_case`` / ``parse_error_details`` 仍从 ``manual_tests.excel.e2e_case`` 取——
后者是**有意为之**，用以守护上提后保留的再导出没有断（21 个 per-feature 脚本仍走该路径）。
"""

import dataclasses

from manual_tests.error_case import (
    PENDING_ADDIN_92,
    ErrorCaseFields,
    ErrorCaseVerdict,
    judge_error_case,
    verdict_note,
)
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

    def test_message_containing_details_literal(self) -> None:
        """对抗用例：wire message 自身含 '(details: {...})' 字面量 → 仍解析尾部真 details"""
        wire = {
            "code": "3010",
            "message": 'weird (details: {"fake": 1}) in message',
            "details": {"kind": "chart"},
        }
        assert parse_error_details(format_wire_error(wire)) == {"kind": "chart"}


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


class TestJudgeErrorCase:
    """judge_error_case() — xfail 三态编排（#90 上提为 Word/PPT/Excel 共享）"""

    def test_strict_hit_without_xfail_is_pass(self) -> None:
        verdict, msg = judge_error_case(False, "3012: No match", "3012")
        assert verdict.passed is True
        assert "✅" in msg

    def test_strict_hit_with_xfail_is_xpass(self) -> None:
        """Add-In 已接线 → 提示可摘标转正（否则摘标时机无从发现）"""
        verdict, msg = judge_error_case(False, "3012: No match", "3012", xfail_reason=PENDING_ADDIN_92)
        assert verdict.passed is True
        assert "XPASS" in msg
        assert "可摘" in msg

    def test_code_miss_with_xfail_is_xfail(self) -> None:
        """Add-In 未接线仍返 3000 → 计通过但标 ⚠️ XFAIL，套件不红"""
        verdict, msg = judge_error_case(False, "3000: Document error", "3012", xfail_reason=PENDING_ADDIN_92)
        assert verdict.passed is True
        assert "XFAIL" in msg

    def test_code_miss_without_xfail_fails(self) -> None:
        verdict, msg = judge_error_case(False, "3000: Document error", "3012")
        assert verdict.passed is False
        assert "❌" in msg

    def test_unexpected_success_fails_even_with_xfail(self) -> None:
        """**关键**：xfail 只赦免「码不对」，不赦免「本应失败却成功」——后者是真回归"""
        verdict, msg = judge_error_case(True, None, "3012", xfail_reason=PENDING_ADDIN_92)
        assert verdict.passed is False
        assert "❌" in msg

    def test_details_mismatch_with_xfail_is_xfail(self) -> None:
        """码对但 details 不符，且响应确为失败 → 仍按 XFAIL 兜底。

        实收值取 ``element`` 而非编造的词表外 token——两者都在 error-handling.md 的
        ``details.kind`` 词表内，更贴近「Add-In 判成了另一类对象」的真实错法。
        """
        err = '3008: Out of range (details: {"kind": "element"})'
        verdict, msg = judge_error_case(False, err, "3008", {"kind": "slide"}, PENDING_ADDIN_92)
        assert verdict.passed is True
        assert "XFAIL" in msg

    def test_expected_label_includes_details(self) -> None:
        """文案里带上 details 期望，真机排查时不必回查源码"""
        _, msg = judge_error_case(False, "3000: x", "3008", {"kind": "slide"})
        assert "details⊇" in msg
        assert "slide" in msg

    def test_none_error_without_xfail_fails(self) -> None:
        """失败但 error 为 None（连接层异常等）→ 无 xfail 时判失败"""
        verdict, msg = judge_error_case(False, None, "3012")
        assert verdict.passed is False
        assert "❌" in msg

    def test_none_error_with_xfail_is_xfail(self) -> None:
        """失败且 error 为 None → 有 xfail 时仍按 XFAIL 兜底（响应确为失败即满足前提）"""
        verdict, msg = judge_error_case(False, None, "3012", xfail_reason=PENDING_ADDIN_92)
        assert verdict.passed is True
        assert "XFAIL" in msg


class TestVerdictIsMachineReadable:
    """调用方必须能从**返回值**区分四态，而不是去 match 展示文案（#90 复审 🔴）。

    事故复盘：``test_excel_e2e.expect_error`` 曾用 ``"✅" in message`` 反推「严格命中」，
    但该调用点恒传 ``xfail_reason``——严格命中时返回的是 🎉 XPASS 文案（**不含 ✅**），
    条件恒 False，XPASS 被误记为 XFAIL，摘标信号在 UAT 汇总里静默丢失。
    原有 9 条单测全在做文案子串匹配，故未能拦住。
    """

    STRICT_HIT = '3010: not found (details: {"kind": "worksheet"})'

    def test_four_verdicts_are_distinct(self) -> None:
        """四态两两可分——这是「XPASS 不被吞进 XFAIL」的机器可读前提"""
        v_pass, _ = judge_error_case(False, self.STRICT_HIT, "3010")
        v_xpass, _ = judge_error_case(False, self.STRICT_HIT, "3010", xfail_reason=PENDING_ADDIN_92)
        v_xfail, _ = judge_error_case(False, "3000: boom", "3010", xfail_reason=PENDING_ADDIN_92)
        v_fail, _ = judge_error_case(True, None, "3010", xfail_reason=PENDING_ADDIN_92)

        assert (v_pass, v_xpass, v_xfail, v_fail) == (
            ErrorCaseVerdict.PASS,
            ErrorCaseVerdict.XPASS,
            ErrorCaseVerdict.XFAIL,
            ErrorCaseVerdict.FAIL,
        )
        assert len({v_pass, v_xpass, v_xfail, v_fail}) == 4

    def test_passed_maps_to_verdict(self) -> None:
        """passed 与 verdict 不得漂移：有且仅有 FAIL 不通过。

        **逐个列举而非复用实现里的表达式**——写成 ``v.passed is (v is not FAIL)`` 只是把
        实现抄了一遍，实现怎么改断言就怎么跟着变，等于没有独立声明意图。
        """
        assert ErrorCaseVerdict.PASS.passed is True
        assert ErrorCaseVerdict.XPASS.passed is True, "XPASS 是「已接线可摘标」，必须计通过"
        assert ErrorCaseVerdict.XFAIL.passed is True, "XFAIL 是过渡期兜底，计通过、套件不红"
        assert ErrorCaseVerdict.FAIL.passed is False
        # 防新增枚举成员时漏定义通过性
        assert {v.name for v in ErrorCaseVerdict} == {"PASS", "XPASS", "XFAIL", "FAIL"}

    def test_message_must_not_be_used_to_infer_state(self) -> None:
        """钉死事故本身：严格命中 + xfail 时文案**不含** ✅，故文案匹配法必然失效"""
        verdict, message = judge_error_case(False, self.STRICT_HIT, "3010", xfail_reason=PENDING_ADDIN_92)
        assert verdict is ErrorCaseVerdict.XPASS
        assert "✅" not in message
        assert "XPASS" in message

    def test_verdict_note_keeps_xpass_distinguishable(self) -> None:
        """汇总表备注里 XPASS 必须与 XFAIL 可分——否则等于没有摘标探测"""
        note_pass = verdict_note(ErrorCaseVerdict.PASS, False, "x")
        note_xpass = verdict_note(ErrorCaseVerdict.XPASS, False, "x")
        note_xfail = verdict_note(ErrorCaseVerdict.XFAIL, False, "x")
        note_fail = verdict_note(ErrorCaseVerdict.FAIL, True, None)

        assert note_pass == ""
        assert "XPASS" in note_xpass and "可摘" in note_xpass
        assert "XFAIL" in note_xfail
        assert note_xpass != note_xfail
        assert "ok=True" in note_fail


class TestErrorCaseFieldsMixin:
    """ErrorCaseFields —— 三个命名空间的 case dataclass 必须共享同一组错误码字段（#90）。

    防回归点：后续新增 runner / case 类型时漏继承 mixin，会导致该命名空间的错误码用例
    声明了字段却不生效（dataclass 静默忽略未声明字段会直接 TypeError，但漏继承则是
    「看起来能写、判定却读不到」）。
    """

    def test_all_case_types_expose_error_fields(self) -> None:
        from manual_tests.e2e_base import TestCase as WordCase
        from manual_tests.error_case import ErrorCaseFields
        from manual_tests.excel.e2e_case import ExcelCase
        from manual_tests.ppt.e2e_base import PptTestCase

        for case_cls in (WordCase, PptTestCase, ExcelCase):
            assert issubclass(case_cls, ErrorCaseFields), f"{case_cls.__name__} 未继承 ErrorCaseFields"
            fields = {f.name for f in dataclasses.fields(case_cls)}
            assert {"expect_error_code", "expect_error_details", "xfail_reason"} <= fields

    def test_error_fields_are_keyword_only(self) -> None:
        """kw_only 是 mixin 可行的前提——带默认值的基类字段否则会挡住子类必填字段"""
        for f in dataclasses.fields(ErrorCaseFields):
            assert f.kw_only is True, f"{f.name} 必须是 keyword-only"

    def test_subclass_keeps_positional_required_fields(self) -> None:
        """继承 mixin 后，子类的必填位置参数仍可按位置传（不被基类默认值挤掉）"""
        from manual_tests.ppt.e2e_base import PptTestCase

        case = PptTestCase("名称", "fixture.pptx", "描述", expect_error_code="3008")
        assert case.name == "名称"
        assert case.expect_error_code == "3008"
        assert case.xfail_reason is None
