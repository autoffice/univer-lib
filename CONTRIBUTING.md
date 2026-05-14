# Contributing to univer-lib

感谢愿意参与 univer-lib！在提交 issue 或 PR 之前，请花几分钟读完本指南。

Thanks for your interest in contributing to univer-lib! Please take a few minutes to read this guide before opening issues or PRs.

---

## 前置准备 / Prerequisites

- **JDK 8+**（库本身保持 JDK 8 兼容，CI 同时跑 8/11/17/21）
- 使用仓库自带的 `mvnw` / `mvnw.cmd`，**不要**调用系统 `mvn`
- 推荐 IDE 安装 **Lombok** 插件（项目根有 `lombok.config`）

## 本地构建 / Local build

```bash
./mvnw test                                 # 跑全部单测
./mvnw test -Dtest=StyleConverterTest       # 单个测试类
./mvnw -Pcoverage verify                    # JaCoCo 覆盖率报告：target/site/jacoco/index.html
./mvnw -Plint verify                        # Alibaba p3c-pmd 命名检查
./mvnw install -DskipTests                  # 装到 ~/.m2，让 example/backend 解析依赖
```

`example/` 不是 Maven submodule，是独立 Spring Boot 2.7 + Vue 3 演示：

```bash
cd example/backend && ../../mvnw spring-boot:run
cd example/frontend && npm install && npm run dev   # http://localhost:5173
```

## 设计文档先行 / Read the design doc first

权威设计在 `docs/design.md`，public API、跨层改动、字段映射有疑问时以它为准。改动这些前先看它，避免破坏既有约定。

The authoritative design lives in `docs/design.md`. Read it before changing public API, doing cross-layer refactors, or touching field mappings.

## 架构纪律 / Architecture discipline

- **分层**：`io → converter → resource → model / util`。每层只依赖下层；**POI 类型不得泄漏到 `converter` 之上**。
- **公共入口**只有 `io.github.autoffice.univer.UniverXlsx`。其他类是实现细节，不要暴露 POI 类型或内部 converter。
- **Sidecar 模式**：xlsx 表达不出的 Univer 字段（`resources`、`custom`、`appVersion`、`scrollTop/Left` 等）通过 `/univer/metadata.json` 自定义 OPC 部件持久化。修改读写流程前理解 `WorkbookConverter.fromXlsx` 与 `WorkbookConverter.mergeSheetData`。
- **POJO 约定**：全部继承 `AbstractUniverModel`（通过 `extras` 字段做前向兼容）；字段名镜像 Univer TypeScript（`v`/`s`/`t`/`p`/`f`/`si` …）；Lombok 标准配方 `@Data + @Accessors(chain=true) + @NoArgsConstructor + @EqualsAndHashCode(callSuper=true)`，子类**必须**带 `callSuper=true`。
- **Jackson**：序列化反序列化 `IWorkbookData` **只用 `JsonMapper.get()`**，它注册了 `IntegerKeyDeserializer`，Spring Boot 自动配置的 `ObjectMapper` 会在 `cellData` 上挂掉。
- **样式去重**：`StyleConverter` 缓存 `XSSFCellStyle`，绕过 POI 64K 上限。**不要**在 converter 里直接 `wb.createCellStyle()`，统一走 `StyleConverter.toPoiStyle`。
- **共享公式**：`f`+`si` 走 `SharedFormulaRegistry`，`CellConverter` 不处理 `si`。

## TDD 流程 / Test-driven workflow

1. 先写失败测试 → 跑（红）→ 最小实现 → 跑（绿）→ 重构 → 提交。
2. 测试命名：`should_xxx_when_yyy`。
3. 圆环回往测试用 AssertJ `recursive comparison` 并忽略 `extras`。
4. 创建 workbook 一律 `try (XSSFWorkbook wb = new XSSFWorkbook())`，禁止泄漏资源。
5. 提交前必须本地 `./mvnw test` 全绿。

## 提交规范 / Commit style

遵循 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/v1.0.0/) 风格：

```
feat(converter): add hyperlink bridge
fix(style): preserve overline through round-trip
docs: update sidecar diagram
refactor(pivot): drop UUID from pivotId
test+example: cover all 9 sheet features
```

允许的 type：`feat` / `fix` / `refactor` / `test` / `docs` / `chore` / `build` / `perf` / `ci`。

## PR Checklist

提 PR 前请确认：

- [ ] 关联了一个 issue（重大改动至少先开 issue 讨论）
- [ ] `./mvnw test` 全绿
- [ ] `./mvnw -Plint verify` 通过（如果改了源码命名）
- [ ] 新增/修改的 public API 带**中英双语 Javadoc**
- [ ] 改动如果涉及字段映射，更新了 `docs/design.md`
- [ ] 改动量大（5+ 文件 / 跨层）请先在 issue 沟通方案

## 开源协议 / License

本项目以 [Apache-2.0](LICENSE) 协议开源。提交 PR 即表示你同意以同样协议授权你的贡献。

By submitting a pull request you agree to license your contribution under the Apache-2.0 License.

## 联系 / Contact

- 提 Bug / 建议 / 使用问题：[GitHub Issues](https://github.com/autoffice/univer-lib/issues)
- 邮件：hello.aldis@qq.com
