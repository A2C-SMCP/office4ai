"""Excel Worksheet 管理 E2E 测试（#32）。

覆盖 #21 Worksheet 切片 5 个事件：``excel:get:worksheets`` / ``excel:add:worksheet`` /
``excel:delete:worksheet`` / ``excel:rename:worksheet`` / ``excel:activate:worksheet``。

双重验证以 openpyxl ``sheet_names``（增/删/改名）+ ``wb.active``（激活）为主——
``delete:worksheet`` / ``activate:worksheet`` 的 AddIn handler **返回 void**（office4ai DTO
声明的 ``{deleted}`` / ``{activated}`` 字段真机不出现），故不能断言响应体，只能读盘核对。
"""
