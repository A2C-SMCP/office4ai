"""Excel Chart 操作 E2E 测试（#34）。

覆盖 #23 Chart 切片 4 个事件：``excel:insert:chart`` / ``excel:get:charts`` /
``excel:update:chart`` / ``excel:delete:chart``。

图表为视觉对象，以**协议层 get:charts 回读**为主验证（openpyxl chart_count 仅 best-effort）。
update/delete 因依赖 Excel 自动生成的图表名，用 ExcelCase.flow 多步流：先 insert 拿 name，
再操作，最后 get:charts 回读核对。
"""
