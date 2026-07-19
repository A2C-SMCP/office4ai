"""错误码用例判定 —— 三命名空间共享（issue #82 引入 /excel，#90 上提为共享）。

manual e2e 的错误码断言有一个绕不开的过渡期问题：**协议已定案、Add-In 尚未接线**。
此时严格断言必然不满足，但把用例判失败会让整套 e2e 长期挂红、失去信噪比；直接放宽断言
又会在 Add-In 接线后**永远发现不了**它已经接好。

故引入 xfail 三态（本模块 :func:`judge_error_case`）：

* 严格命中 + 无 ``xfail_reason``  → ✅ 正常通过
* 严格命中 + 有 ``xfail_reason``  → 🎉 **XPASS**，提示「可摘标转正」
* 未命中 + 有 ``xfail_reason`` 且确为失败响应 → ⚠️ **XFAIL**，计通过、套件不红
* 其余（含**本应失败却成功**）→ ❌ 失败

最后一条是关键：XFAIL **仍要求响应必须失败**，否则「本应失败却成功」这类真回归会被掩盖。

反解析用的 :func:`parse_error_details` 住在生产侧（``office_workspace``），与其逆函数
``format_wire_error`` 同居一处以杜绝格式漂移——本模块转发导入，供 e2e 侧直接使用。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from office4ai.environment.workspace.office_workspace import parse_error_details

__all__ = [
    "PENDING_ADDIN_80",
    "PENDING_ADDIN_92",
    "ErrorCaseFields",
    "ErrorCaseVerdict",
    "evaluate_error_case",
    "judge_error_case",
    "parse_error_details",
    "verdict_note",
]


class ErrorCaseVerdict(Enum):
    """错误码用例的四态判定。

    调用方**读本枚举**决定如何记录，不要去 match 展示文案里的 emoji——展示层不是机器可读的
    判定结果（见 :func:`judge_error_case` docstring 记录的事故）。

    Attributes:
        PASS: 严格命中且无 xfail 标 —— 正常通过。
        XPASS: 严格命中但仍挂着 xfail 标 —— **Add-In 已接线，可摘标转正**。这是整套 xfail
            机制存在的理由：探测过渡期何时结束。必须与 XFAIL 可区分。
        XFAIL: 未严格命中但确为失败响应，且有 xfail 标 —— 过渡期兜底，计通过、套件不红。
        FAIL: 其余，含**本应失败却成功**（xfail 不赦免此项，那是真回归）。
    """

    PASS = "pass"
    XPASS = "xpass"
    XFAIL = "xfail"
    FAIL = "fail"

    @property
    def passed(self) -> bool:
        """是否计入通过。唯一不通过的是 :attr:`FAIL`。"""
        return self is not ErrorCaseVerdict.FAIL

#: oasp#17 /excel 错误码收敛——Add-In 侧接线跟踪 office-editor4ai#80。
#: 接线后此常量的**全部**引用一并摘除即转正（grep 本名可枚举，勿再写字面量）。
PENDING_ADDIN_80 = "待 Add-In 接线 office-editor4ai#80"

#: oasp#23 四个 3xxx 孤儿码裁决（3012 收敛 / 3007·3008 接线）——跟踪 office-editor4ai#92。
PENDING_ADDIN_92 = "待 Add-In 接线 office-editor4ai#92"


@dataclass
class ErrorCaseFields:
    """错误码路径三件套，供三个命名空间的 case dataclass 继承（#90）。

    ``kw_only=True`` 是关键：dataclass 默认把基类字段排在子类字段之前，带默认值的基类字段
    会挡住子类的必填字段（``name`` 等）而报 TypeError；标记为 keyword-only 后其不参与位置
    参数排序，混入即可。

    Attributes:
        expect_error_code: 置位时**预期失败**且 error 含该权威码（如 ``"3012"``）。比布尔的
            ``expect_failure`` 严格：后者只问「失没失败」，不校验失败得对不对。
        expect_error_details: details **子集**断言（如 ``{"kind": "slide"}``），从错误字符串
            尾部的 ``(details: {...})`` 反解析后逐键比对。子集而非全等——实现补字段不该打红。
        xfail_reason: 置位时启用 xfail 三态，用于「协议已定案但 Add-In 尚未接线」的过渡期。
            取值用本模块的 ``PENDING_ADDIN_*`` 常量，**勿写字面量**（摘标时靠 grep 常量名枚举）。
    """

    expect_error_code: str | None = field(default=None, kw_only=True)
    expect_error_details: dict[str, Any] | None = field(default=None, kw_only=True)
    xfail_reason: str | None = field(default=None, kw_only=True)


def evaluate_error_case(
    ok: bool,
    err: str | None,
    expect_code: str,
    expect_details: dict[str, Any] | None = None,
) -> bool:
    """错误码路径的**严格**判定：必须失败 + error 含权威码 + details 子集匹配。

    details 用**子集**语义而非全等——Add-In 可能回带规范未要求的附加字段，断言只钉住
    协议明确要求的键，避免实现补字段就把用例打红。
    """
    if ok or expect_code not in (err or ""):
        return False
    if expect_details:
        actual = parse_error_details(err or "")
        if actual is None:
            return False
        return all(actual.get(key) == value for key, value in expect_details.items())
    return True


def judge_error_case(
    ok: bool,
    err: str | None,
    expect_code: str,
    expect_details: dict[str, Any] | None = None,
    xfail_reason: str | None = None,
) -> tuple[ErrorCaseVerdict, str]:
    """错误码路径四态判定，返回 ``(判定, 展示文案)``。

    把「严格判定 + xfail 三态 + 文案」收在一处，供 Word / PPT / Excel 三个 runner 共用——
    否则每个 runner 都要抄一遍分支，摘标时改三处。文案里的 ✅/🎉/⚠️/❌ 表达的是**用例判定**
    而非正负向（负例按预期失败即 ✅）。

    ⚠️ 返回的 ``message`` 是**展示文案**，**不得**用于反推判定状态（如 ``"✅" in message``）——
    判定一律读 :class:`ErrorCaseVerdict`。曾出过一次事故：调用方用 ``"✅" in message`` 判「严格
    命中」，但该调用点恒传 ``xfail_reason``，严格命中时返回的是 🎉 XPASS 文案（不含 ✅），
    导致条件恒 False、XPASS 被误记为 XFAIL，**摘标信号在汇总里静默丢失**。
    """
    expected = expect_code + (f" details⊇{expect_details}" if expect_details else "")

    if evaluate_error_case(ok, err, expect_code, expect_details):
        if xfail_reason:
            msg = f"   🎉 XPASS [{expected}] → err={err}（Add-In 已接线，可摘 xfail_reason 转正）"
            return ErrorCaseVerdict.XPASS, msg
        return ErrorCaseVerdict.PASS, f"   ✅ 预期失败 [{expected}] → ok={ok} err={err}"

    # XFAIL 仅赦免「码不对」，不赦免「本应失败却成功」——后者是真回归，必须打红
    if xfail_reason and not ok:
        return ErrorCaseVerdict.XFAIL, f"   ⚠️  XFAIL（{xfail_reason}）期望 [{expected}] 实收: err={err}"

    return ErrorCaseVerdict.FAIL, f"   ❌ 预期失败 [{expected}] → ok={ok} err={err}"


def verdict_note(verdict: ErrorCaseVerdict, ok: bool, err: str | None) -> str:
    """把判定压成一行备注，供 UAT 汇总表（如 ``ResultCollector.record``）消费。

    与 :func:`judge_error_case` 的展示文案分开：那条给人看，这条进汇总表。**XPASS 必须与
    XFAIL 可区分**——XPASS 是「Add-In 已接线、可摘标」的唯一探测信号，混进 XFAIL 就等于
    永远发现不了摘标时机，那样这套 xfail 机制就只剩掩盖作用了。
    """
    if verdict is ErrorCaseVerdict.PASS:
        return ""
    if verdict is ErrorCaseVerdict.XPASS:
        return f"XPASS 可摘 xfail 标: {err}"
    if verdict is ErrorCaseVerdict.XFAIL:
        return f"XFAIL 实际: {err}"
    return f"实际: ok={ok} err={err}"
