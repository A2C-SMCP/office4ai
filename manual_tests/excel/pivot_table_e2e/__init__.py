"""Excel PivotTable 操作 E2E 测试（#35）。

覆盖 #24 PivotTable 切片 3 个事件：``excel:insert:pivotTable`` /
``excel:get:pivotTables`` / ``excel:delete:pivotTable``。

透视表为视觉对象，以**协议层 get:pivotTables 回读**为唯一验证依据（openpyxl 透视支持有限）。
仅用 sourceAddress + targetAddress 创建空透视表骨架，不构造 rows/columns/values/filters。
delete 用 ExcelCase.flow 多步流（insert→delete→get 回读核对消失）。
"""
