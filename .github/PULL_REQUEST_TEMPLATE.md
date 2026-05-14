## 关联的 Issue / Related issue

> Closes #xxx 或 Refs #xxx

## 改动说明 / What & Why

> 这个 PR 解决了什么问题？为什么这样改？

## 自检清单 / Self-checklist

- [ ] `./mvnw test` 全绿
- [ ] 如果改了源码命名：`./mvnw -Plint verify` 通过
- [ ] 新增 / 修改的 public API 带**中英双语 Javadoc**
- [ ] 改动涉及字段映射或公开协议时，已同步更新 `docs/design.md`
- [ ] 改动若是新功能或修 bug，已经新增对应单测（命名 `should_xxx_when_yyy`）
- [ ] commit message 遵循 Conventional Commits（`feat(xxx): ...` / `fix(xxx): ...`）
- [ ] 如果是 5+ 文件 / 跨层改动，已经先在 issue 与维护者对齐方案

## 测试方式 / How to verify

> 审核者怎么验证这个 PR？提供测试文件、命令或截图。

## 风险与影响范围 / Risk & blast radius

> 是否影响 public API？是否破坏既有 round-trip？是否需要在 CHANGELOG 标 BREAKING？
