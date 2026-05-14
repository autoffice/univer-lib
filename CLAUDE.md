# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Communication

总是用中文回答我，技术词汇（类名、方法名、库名、命令、配置项等）保持英文。

## Repository purpose

Java library `io.github.autoffice:univer-lib` for high-fidelity bidirectional conversion between xlsx files and Univer Sheets `IWorkbookData` JSON snapshots. JDK 8 compatible.

设计文档（权威）：`docs/design.md`。public API、跨层改动、字段映射有疑问时以它为准。

## Common commands

All Maven commands use the wrapper. Do **not** invoke system `mvn`.

### 测试 / Testing

```bash
./mvnw test                                                                # 全部单测
./mvnw test -Dtest=StyleConverterTest                                      # 单个测试类
./mvnw test -Dtest=StyleConverterTest#should_roundtrip_font_and_alignment  # 单个方法
./mvnw test -Dtest='ConditionalFormattingConverterTest,StyleConverterTest' # 多个类
./mvnw -Pcoverage verify                                                   # JaCoCo 报告：target/site/jacoco/index.html
```

### 静态检查 / Lint

```bash
./mvnw -Plint verify                     # Alibaba p3c-pmd（PMD 6 + maven-pmd-plugin 3.21.2 锁死）
./mvnw -DskipTests -Plint verify         # 只跑 lint，不跑测试，速度更快
```

### 构建与本地安装 / Build & local install

```bash
./mvnw clean                             # 清理 target/ 与 .flattened-pom.xml
./mvnw -DskipTests package               # 构建 jar（仍会跑 flatten-maven-plugin）
./mvnw install -DskipTests               # 装到 ~/.m2/repository，让 example/backend 解析依赖
./mvnw dependency:tree                   # 排查依赖冲突
```

> 版本号由 `${revision}` 占位符 + `flatten-maven-plugin` 管理。临时改版本不要直接改 `<revision>`，用 `-Drevision=1.2.3` 覆盖。

### 发布 / Release（仅 CI / 维护者）

`release` profile 启用 source jar、javadoc jar、GPG 签名、license 头检查、Sonatype Central 发布。**日常开发不要用**（GPG 不可用会失败）。

```bash
./mvnw -Prelease -DskipTests verify      # 本地 dry-run：会跑 license:check 和 GPG sign
./mvnw -Prelease deploy                  # 真正发布到 Maven Central（CI 会做，本地一般不需要）
```

### License 头管理 / License headers

新文件创建后必须带 Apache-2.0 头，CI 在 `release` profile 里会校验。

```bash
./mvnw -Prelease com.mycila:license-maven-plugin:4.6:format   # 自动给所有 Java 文件加头
./mvnw -Prelease com.mycila:license-maven-plugin:4.6:check    # 验证所有文件都有头
```

### Example demo（独立子项目，非 Maven submodule）

```bash
cd example/backend && ../../mvnw spring-boot:run     # JDK 8+ 即可
cd example/frontend && npm install && npm run dev    # http://localhost:5173
```

## Architecture (big picture)

The library is layered. Each layer only depends on the layer below it; **POI types never leak above the `converter` package**.

```
io                  UniverXlsxReader / UniverXlsxWriter   (public-facing thin wrappers)
  ↓
converter           WorkbookConverter
                      ├─ WorksheetConverter
                      │    ├─ CellConverter (uses StyleConverter)
                      │    ├─ RichTextConverter
                      │    └─ SharedFormulaRegistry
                      └─ StyleConverter
  ↓
resource            SidecarPart  (custom OPC part /univer/metadata.json)
  ↓
model               IWorkbookData / IWorksheetData / ICellData / IStyleData / ...   (POJOs)
util                JsonMapper / ColorUtils / LengthUtils / IntegerKeyDeserializer
```

The single public entry point is `io.github.autoffice.univer.UniverXlsx` (static `read`/`write` methods + `UniverXlsxOptions` builder). Everything else is implementation detail; do not expose POI types or internal converters.

### The sidecar pattern (critical to understand)

xlsx cannot natively represent every Univer field (`resources`, `custom`, `appVersion`, `IStyleData.pd` padding, overline, `scrollTop/Left`, etc.). The library writes the **complete** `IWorkbookData` JSON into a custom OPC part `/univer/metadata.json` inside the xlsx. On read:

1. The sidecar (if present) is loaded as the baseline `IWorkbookData`.
2. The xlsx is parsed normally; xlsx-derived values overwrite **content** fields (`cellData`, `mergeData`, `rowData`, `columnData`, freeze, etc.) on the sidecar baseline.
3. Sidecar-only fields (`resources`, `custom`, `appVersion`, ...) survive untouched.

This is what makes round-trip lossless. External xlsx files (no sidecar) read fine — fields without xlsx representation are simply absent. See `WorkbookConverter.fromXlsx` and `WorkbookConverter.mergeSheetData`.

### POJO conventions

- All POJOs extend `AbstractUniverModel`, which captures unknown JSON fields via `extras` (`@JsonAnyGetter`/`@JsonAnySetter`). This guarantees forward compatibility with future Univer versions — never break this contract.
- Field **names mirror Univer TypeScript** (`v`, `s`, `t`, `p`, `f`, `si`, `bl`, `fs`, ...). Lombok `@Data + @Accessors(chain=true) + @NoArgsConstructor + @EqualsAndHashCode(callSuper=true)` is the standard recipe; subclasses **must** include `callSuper=true` for equality to work, since `AbstractUniverModel.equals` compares `extras`.
- Boolean-as-number fields use the `BooleanNumber` enum (serializes to `0`/`1`).
- Public classes/methods carry **bilingual (Chinese + English) Javadoc** per Alibaba spec. Keep this when adding code.

### Jackson setup

Always use `JsonMapper.get()` for serializing/deserializing `IWorkbookData`. It registers `IntegerKeyDeserializer` globally so that `Map<Integer, Map<Integer, ICellData>>` (cellData) and `Map<Integer, IRowData>` deserialize correctly from JSON's stringified numeric keys. Spring Boot's auto-configured `ObjectMapper` will fail on `cellData` parsing — the `example/backend` controller deliberately bypasses it via `JsonMapper.get()`.

### Style deduplication

`StyleConverter` keeps a `Map<styleId, XSSFCellStyle>` cache keyed by a stable hash of the `IStyleData` JSON. Two equal `IStyleData` instances produce the *same* `XSSFCellStyle` object — this avoids POI's 64K cell-style limit. Don't bypass the cache by calling `wb.createCellStyle()` directly in converters; go through `StyleConverter.toPoiStyle`.

### Shared formulas (`si`)

`SharedFormulaRegistry` groups cells by formula `si`. On read: same expression → same generated `si` (UUID), master cell (the canonical `f` holder) is bottom-right of the group. On write: registered cells are post-processed in `applyOnWorkbook` so the master sits bottom-right per Univer's contract. `WorksheetConverter.writeCellData` defers `f`+`si` cells to the registry; `CellConverter` does NOT handle `si`.

## Testing approach

- TDD: write failing test → run → minimal impl → run → commit. The plan steps follow this; keep doing it for new features.
- Test naming: `should_xxx_when_yyy`.
- 70 tests at the time of writing; coverage ≈47% line. POI fallback branches are intentionally uncovered.
- Round-trip tests use AssertJ recursive comparison with `extras` ignored.
- Use `try (XSSFWorkbook wb = new XSSFWorkbook())` in tests (and converters that create a workbook); never leak.

## Things to be careful about

- `mvnw` is committed in the repo. Use it; do not regenerate.
- POI 5.2.5 uses `setQuotePrefixed`/`getQuotePrefixed` (NOT `setQuotePrefix`). Force-text cells rely on this — see `CellConverter`.
- Length conversions in `LengthUtils` are **approximations** (1 inch = 96 px = 72 pt; 1 char ≈ 7 px). Documented as such; don't "fix" them without checking the spec.
- `IWorkbookData.resources` is intentionally typed as `Object` — opaque JSON pass-through for plugin payloads.
- `strictMode` in `UniverXlsxOptions` is currently a reserved switch; no converter throws `UniverXlsxUnsupportedFeatureException` yet. README documents this gap.
- The `lint` profile uses Alibaba `p3c-pmd:2.1.1` which only works with PMD 6.x; the pom pins `maven-pmd-plugin:3.21.2` (which uses PMD 6) for this reason. Don't bump it to 3.22+ until p3c-pmd ships a PMD 7 release.

## Workflow notes

设计决策的源头是 `docs/design.md`。在改动 public API 或做跨层变更之前先看它，避免破坏既有约定（分层、sidecar、POJO/Jackson 约定、样式去重、共享公式）。
