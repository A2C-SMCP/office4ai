"""Excel Find & Filter E2E 测试（#36）。

覆盖 #25 Find&Filter 切片 3 个事件：``excel:find:values`` / ``excel:set:autoFilter`` /
``excel:clear:autoFilter``。

find:values 验证 matchCase / matchEntireCell（默认 false）与无命中；autoFilter 以 openpyxl
``auto_filter_ref`` 双重验证（set 后 ref 落地、clear 后 ref 清空）。
"""
