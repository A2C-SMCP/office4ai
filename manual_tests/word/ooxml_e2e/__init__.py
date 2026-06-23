"""
Word OOXML Round-Trip End-to-End Tests

测试 word:get:ooxml / word:insert:ooxml 事件对（OOXML 片段 round-trip）。

⚠️ 外部依赖：真机运行需 office-editor4ai (word-editor4ai) Add-In 已实现
   word:get:ooxml / word:insert:ooxml 两个 handler（经 cross-ask 已锁定契约，
   见 office4ai#46）。Add-In 就绪前，真机执行会在 get 阶段超时/报 3016。
"""
