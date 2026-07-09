---
name: demo-authoring
description: Demo authoring SKILL fixture — exercises the skill:// resources source mode (root + scripts + references + binary asset). Triggers in S3 producer tests only.
version: 0.1.0
license: MIT
---

# Demo Authoring SKILL (fixture)

中文：本 SKILL 仅用于 office4ai S3（issue #59）的 `skill://` producer 侧测试，
不是生产能力包。它示范 SKILL 包如何携带 `scripts/`（参考脚本）、`references/`（指导文档）
与二进制 `assets/`（模板资产），并验证 `resources` source 模式的物化契约。

English: This SKILL exists only for office4ai S3 (issue #59) `skill://` producer-side tests.
It demonstrates a SKILL package carrying `scripts/`, `references/`, and a binary `assets/`
payload, validating the `resources` source-mode materialization contract.

## 用法约定（示意）

LLM 写脚本时 **import** helper（`from office4ai.office.authoring.helpers import ...`），
读 `scripts/gen_demo.py` 学模式后仿写自己的几行 glue，提交给 `office_run_script` —— **不抄** helper 源码。
