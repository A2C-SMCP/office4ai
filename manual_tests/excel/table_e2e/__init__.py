"""Excel Table 操作 E2E 测试（#33）。

覆盖 #22 Table 切片 6 个事件：``excel:insert:table`` / ``excel:get:table`` /
``excel:get:tables`` / ``excel:add:tableRow`` / ``excel:delete:tableRow`` / ``excel:sort:table``。

双重验证以 openpyxl ``table_names`` / 单元格值读盘为主。错误码两条路径：
- 表不存在 → ``3010`` ELEMENT_NOT_FOUND(kind:table)；rowIndex 越界（合法非负）→ ``4004``
  PARAM_OUT_OF_RANGE（oasp#17 定案；Add-In 接线前 XFAIL）；
- tableId 空串 / rowIndex·columnIndex 负数（Zod nonnegative 失败）→ ``4000`` VALIDATION_ERROR。
"""
