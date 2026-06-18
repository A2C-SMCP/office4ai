"""find_filter_e2e 夹具构建器。

生成单个夹具 ``filt.xlsx``（提交入库，供真机打开），覆盖 #36 全部 find/filter 用例：
查找（find:values）/ 自动筛选（set/clear:autoFilter）。

一张表：

- ``Data`` —— **activeSheet**，A1:C5 网格，**Product 列大小写混合**（Apple/apple）以验证
  ``matchCase``，并为 ``set:autoFilter`` 提供可筛选的 Region/Product 文本列。

数据布局::

    Region   Product  Amount
    East     Apple    100
    West     apple    200
    East     Banana   300
    South    apple    400

- find "apple" 默认（不区分大小写、子串）→ 命中 B2/B3/B5 共 3 个
- find "apple" matchCase=true → 仅小写 B3/B5 共 2 个
- find "App" matchEntireCell=true → 无单元格整体等于 "App" → 0 命中
- find "East" 限定 address=A1:A5 → A2/A4 共 2 个

重新生成：
    uv run python manual_tests/excel/find_filter_e2e/_fixtures.py
"""

from __future__ import annotations

from pathlib import Path

FIXTURE_DIR = Path(__file__).parent.parent / "fixtures" / "find_filter_e2e"

_DATA_ROWS: list[list[object]] = [
    ["Region", "Product", "Amount"],
    ["East", "Apple", 100],
    ["West", "apple", 200],
    ["East", "Banana", 300],
    ["South", "apple", 400],
]


def _fill(ws, rows: list[list[object]]) -> None:
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=value)


def build_filt(path: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    data = wb.active
    data.title = "Data"
    _fill(data, _DATA_ROWS)
    wb.active = 0  # activeSheet = Data
    wb.save(str(path))


def build_all() -> dict[str, Path]:
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}
    for name, builder in {"filt.xlsx": build_filt}.items():
        target = FIXTURE_DIR / name
        builder(target)
        paths[name] = target
        print(f"📝 生成夹具: {target}")
    return paths


def ensure_fixtures() -> dict[str, Path]:
    """确保 filt.xlsx 夹具存在（缺失才生成）。"""
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    if (FIXTURE_DIR / "filt.xlsx").exists():
        return {"filt.xlsx": FIXTURE_DIR / "filt.xlsx"}
    return build_all()


if __name__ == "__main__":
    build_all()
