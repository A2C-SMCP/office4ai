# filename: gen_demo.py
"""参考脚本示意（fixture）：演示「import helper 而非抄」的用法约定。

Reference script (fixture): demonstrates the "import the helper, don't inline it" convention.
LLM 读此脚本学模式后仿写自己的 glue，提交给 office4ai 的 office_run_script 在稳定运行时执行。
"""

from __future__ import annotations


def main() -> None:
    # 示意：真实 SKILL 里此处会是
    #   from office4ai.office.authoring.helpers import instantiate_from_template, fill_sdt_controls
    #   instantiate_from_template("template.dotx", "out.docx")
    #   fill_sdt_controls("out.docx", {"title": "Demo"})
    print("demo authoring script — replace with your own glue calling the helpers")


if __name__ == "__main__":
    main()
