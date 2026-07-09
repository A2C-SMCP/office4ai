# office4ai authoring SKILL 包分发根 / authoring SKILL package root

本目录是 office4ai 通过 A2C `skill://` 通道对外分发的 **SKILL 能力包**存放根。
This directory is the root of the authoring **SKILL capability packages** office4ai distributes
over the A2C `skill://` channel.

## 机制（milestone #4 · S3，issue #59）

`OfficeMCPServer._register_resources()` 启动时扫描本目录**一级子目录**，凡含 `SKILL.md` 者经
`office4ai.a2c_smcp.resources.skill.SkillResource`（`resources` source 模式）暴露为
`skill://com.a2c-smcp.office4ai/<skill-name>` 及其子文件资源；A2C Computer 负责物化（staging）到本地、
合成全局 name `mcp:office4ai:<frontmatter.name>`。

- 扫描根可经环境变量 `OFFICE4AI_SKILLS_ROOT` 覆盖（部署/测试用）。
- 契约与设计见 `docs/milestone4_authoring_desktop_spec.md` 与 a2c-smcp-protocol `docs/specification/skill.md`。

## 约定

每个 SKILL = 一个子目录：

```
office/skills/
  <skill-name>/            # 目录名 == SKILL.md frontmatter.name（严格 kebab）
    SKILL.md               # 必需，YAML frontmatter 至少含 name + description
    scripts/               # 可选：参考脚本；LLM 读之学模式，脚本里 import 而非抄
    references/            # 可选：进一步指导文档（渐进式披露）
    assets/                # 可选：二进制模板资产（.dotx/.potx/.xltx 等，producer 以 blob 分发）
```

**S3 地基阶段本目录为空**（除本 README）——不落任何生产 SKILL。生产 SKILL 由后续子任务落入：
- `create-office-file`（S4 / W1，issue #60）
- `edit-office-file`（S5 / W2，issue #61）
- `extract-template`（S6 / W3，issue #62）
