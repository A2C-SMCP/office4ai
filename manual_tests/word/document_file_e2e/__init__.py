"""
Word whole-document .docx Round-Trip End-to-End Tests (#46-D)

测试 word:get:documentFile / word:insert:documentFile 事件对（整篇 base64 .docx round-trip）。
这是 insertOoxml 吃不下整篇 Flat OPC 包后改走的整篇路（base64 .docx + insertFileFromBase64）。

⚠️ 外部依赖：真机运行需 office-editor4ai (word-editor4ai) Add-In 已实现
   word:get:documentFile / word:insert:documentFile 两个 handler（cross-ask 重锁契约，#63）。
⚠️ F2 保真：insertFileFromBase64 是正文替换，文档级部件（页眉页脚/文档级 sectPr/docProps/customXml）
   不保证完整 round-trip——本套显式校验正文文本保真，不假设字节级对称。
"""
