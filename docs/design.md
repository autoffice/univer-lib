# Univer xlsx Converter 设计

> Java 侧在 xlsx 与 Univer `IWorkbookData` 之间进行高保真双向转换的轻量库。

## 1. 目标与范围

- **功能**：`xlsx → IWorkbookData` 与 `IWorkbookData → xlsx` 的对称转换。
- **保真度**：覆盖 `IWorkbookData` 全部字段（含 `styles / mergeData / 行列宽高 / 冻结 / 隐藏 / RTL / 网格线 / zoomRatio / resources / custom / 富文本 p / 共享公式 si`）。
- **运行环境**：JDK 8+。
- **坐标**：`io.github.autoffice:univer-lib`，包名 `io.github.autoffice.univer`。
- **非目标**：公式计算、富文本渲染、解释 `resources` 内部业务语义（仅作不透明 JSON round-trip）、xlsm/xlsb 支持。

## 2. 架构总览

Maven 单模块，按职责分层：

| 包 | 职责 |
|----|------|
| `...univer` | 对外门面 `UniverXlsx`、`UniverXlsxOptions`、异常类型 |
| `...univer.model` | 与 Univer TS 接口一一对应的 POJO |
| `...univer.converter` | POI ↔ POJO 字段级映射（对称设计） |
| `...univer.io` | `UniverXlsxReader` / `UniverXlsxWriter` 内部实现 |
| `...univer.resource` | OPC 边车分区读写（`/univer/metadata.json`） |
| `...univer.util` | 颜色、长度换算、枚举映射、A1 引用、JSON KeyDeserializer |

**依赖**：Apache POI（`poi` + `poi-ooxml`）、Jackson Databind、Lombok（`provided`）、JUL 日志；测试：JUnit 5、AssertJ。

**对外 API（两类入口 + 一份配置）**：
```java
IWorkbookData wb = UniverXlsx.read(inputStream);
UniverXlsx.write(wb, outputStream);
UniverXlsx.read(path, UniverXlsxOptions.builder().strictMode(true).build());
```

## 3. POJO 数据模型

- 字段名与 Univer TS 短名一致（`v/s/t/p/f/si/...`），通过 `@JsonProperty` 绑定；Java 端用 Lombok `@Data + @Accessors(chain=true) + @NoArgsConstructor` 提供可读 setter / getter。
- 布尔语义字段统一用 `BooleanNumber{FALSE(0), TRUE(1)}`，Jackson 序列化为数字。
- `Nullable<T>` 用引用类型 + `@JsonInclude(NON_NULL)`，缺省字段不出现在 JSON。
- `IObjectMatrixPrimitiveType<ICellData>` 映射为 `Map<Integer, Map<Integer, ICellData>>`（`LinkedHashMap` 保序），通过自定义 `IntegerKeyDeserializer` 处理字符串数字键。
- 所有 POJO 继承 `AbstractUniverModel`，内部持有 `Map<String,Object> extras` + `@JsonAnyGetter/@JsonAnySetter`，承接未知字段，保证 Univer 版本升级不丢数据。

**类清单**：`IWorkbookData`、`IWorksheetData`、`ICellData`、`IStyleData`、`IColorStyle`、`ITextDecoration`、`IBorderData`、`IBorderStyleData`、`ITextRotation`、`IPaddingData`、`INumfmtLocal`、`IRange`、`IRowData`、`IColumnData`、`IFreeze`、`IDocumentData`（含 `body.dataStream / textRuns / paragraphs`）。

**枚举**：`CellValueType(STRING=1, NUMBER=2, BOOLEAN=3, FORCE_TEXT=4)`、`HorizontalAlign`、`VerticalAlign`、`WrapStrategy`、`TextDirection`、`LocaleType`、`BorderStyleTypes`。

## 4. xlsx ↔ Univer 映射规则

### 4.1 单元格值与类型

| POI CellType | Univer `t` | Univer `v` 存储形式 |
|--------------|-----------|--------------------|
| STRING       | 1         | 原字符串 |
| NUMERIC      | 2         | `double`（日期格式由 `s.n.pattern` 保留） |
| BOOLEAN      | 3         | `0/1` |
| FORMULA      | 按 cachedValue | `f = cell.formula`，`v = cachedValue`，`t` 按 cachedValue 推断 |
| STRING + quotePrefix | 4 | 原字符串 |

空单元格不写入 `cellData`。

### 4.2 共享公式 `si`

- 读：POI 的 shared formula group 本身无稳定 id，按「公式表达式 + 主格坐标」生成 `si = UUID`，主格写 `f + si`，从属格仅写 `si`。`si` 按 `(sheetIndex, expression)` 分组，跨 sheet 同公式不会冲突。
- 写：同 `si` 分组中主格必须位于右下角（与 Univer 约定一致）；其它格写成 POI shared formula 引用。

### 4.3 富文本 `p`（IDocumentData）

- 读：拆解 `XSSFRichTextString` 的 runs → `p.body.dataStream + textRuns[{st,ed,ts:{ff,fs,bl,it,cl,...}}]`；以 `\n` 分段填充 `paragraphs`。
- 写：反向组装；`v` 与 `p` 并存时，读取仅保留 `p`，写回 xlsx 以 `p.toPlainText()` 作为 inline string。
- 未识别的 `IDocumentData` 扩展字段进入 `extras`，通过边车回写以保证无损。

### 4.4 样式 `s`（IStyleData）

- **去重**：遍历所有 cell 的 `IStyleData`，以稳定 hash 作为 `styleId`，写入 `IWorkbookData.styles`；cell 端只存 `s = styleId`，避免 xlsx 64K style 上限。FORCE_TEXT 用 `styleIdOf + "#qp"` 二级缓存，避免每个文本格式 cell 复制一份样式。
- **字段映射**：

| Univer | POI | 备注 |
|--------|-----|------|
| `ff` | `Font.name` | |
| `fs` | `Font.fontHeightInPoints` | |
| `bl / it` | `Font.bold / italic` | `BooleanNumber` |
| `ul / st` | `Font.underline / strikeout` | |
| `ol` | 无原生 | 边车保留 |
| `cl.rgb` | `XSSFFont.color` | 透明度丢弃 |
| `va` | `Font.typeOffset` | 1=NONE, 2=SUB, 3=SUPER |
| `bg.rgb` | `fillForegroundColor + SOLID_FOREGROUND` | |
| `bd.{t,b,l,r,tl,tr,bl,br}` | POI `BorderStyle + BorderColor` | `s` 映射 0..13 |
| `ht / vt` | `HorizontalAlignment / VerticalAlignment` | |
| `tb` | `WrapText / ShrinkToFit` | 1=overflow, 2=clip, 3=wrap |
| `tr.a / tr.v` | `CellStyle.rotation` | `v=1` → rotation=255 |
| `pd` | indent | 仅 `l/r` 近似映射，其余落边车 |
| `n.pattern` | `DataFormat.getFormat` | |

### 4.5 工作表级字段（IWorksheetData）

| Univer | POI | 备注 |
|--------|-----|------|
| `name / tabColor / hidden` | `Sheet.name / XSSFSheet.tabColor / Workbook.sheetVisibility` | |
| `rowCount / columnCount` | 仅写实际使用范围 | 以边车值优先还原 |
| `defaultColumnWidth / defaultRowHeight` | `Sheet.defaultColumnWidth(字符) / defaultRowHeightInPoints` | Univer 用 px，工具类按 `1ch≈7px, 1pt=96/72 px` 近似换算 |
| `rowData[i].{h,hd}` | `Row.heightInPoints / zeroHeight` | `ah` 落边车 |
| `columnData[j].{w,hd}` | `Sheet.setColumnWidth / setColumnHidden` | |
| `mergeData` | `Sheet.addMergedRegion` | |
| `freeze` | `Sheet.createFreezePane` | `startRow/Column=-1` 表示无冻结 |
| `showGridlines / rightToLeft / zoomRatio` | `Sheet.displayGridlines / XSSFSheet.rightToLeft / Sheet.setZoom` | |
| `scrollTop / scrollLeft / rowHeader / columnHeader / defaultStyle` | 无原生对应 | 全部落边车 |

### 4.6 工作簿级字段

- `id / name / appVersion / locale / resources / styles(完整) / sheetOrder` 一律写入 `/univer/metadata.json` 边车。
- `sheetOrder` 默认按 `workbook.getSheetAt(i)` 顺序还原，边车仅在需要保留额外顺序信息时覆盖。
- 读取时通过 `nameToId`（来自 sidecar）把 POI sheet name 映射回原始 sheet id，避免 sidecar 与 xlsx 同时被写入产生重影 sheet。外部 xlsx（无 sidecar）回退用 sheet name 作为 sid。
- 外部 xlsx 读取时，`IWorkbookData.styles` 由 `StyleConverter.idToStyle` 注册表填充，保证 cell 端的 `s = "<styleId>"` 总有对应的内联样式。

### 4.7 边车文件 `/univer/metadata.json`

- OPC 自定义分区；Content-Type `application/json`；relationship type `http://schemas.autoffice.io/univer/2026/metadata`。
- 内容：完整 `IWorkbookData` JSON（深拷贝），作为**权威源**；xlsx 仅作 fallback。
- 读取优先级：边车命中 → 边车值；边车缺失 → 由 xlsx 推断，`appVersion/locale/id` 等不可推断字段使用默认值。
- 写入失败时抛 `UniverXlsxWriteException`，遵循统一异常体系。

## 5. 数据流与错误处理

### 5.1 读流程

```
InputStream → OPCPackage.open
  ├─ (1) 读 /univer/metadata.json（命中则作为基线 IWorkbookData）
  ├─ (2) XSSFWorkbook 遍历：Workbook 属性 + 每 Sheet 属性 + cellData + mergeData + rowData/columnData + freeze...
  ├─ (3) Style 去重：Reader 维护 Map<IStyleData, styleId>
  └─ (4) 合并：内容字段（cellData / mergeData / rowData / columnData / freeze）以 xlsx 为准；
          辅助字段（resources / custom / appVersion / extras）保留 sidecar；cell 级合并保留 sidecar 独有的 `p`、`custom`
→ IWorkbookData
```

### 5.2 写流程

```
IWorkbookData
  ├─ (1) 规范化：将 styles Map 解引用回每格完整 IStyleData 视图
  ├─ (2) 新建 XSSFWorkbook：按 sheetOrder 建 sheet → 属性 → row/column/cell → 合并区 → freeze
  │      Style 缓存：IStyleData → POI CellStyle/Font，hash 去重
  ├─ (3) 共享公式：按 (sheetIndex, si) 分组，主格置右下
  ├─ (4) 写边车：完整 IWorkbookData JSON
  └─ (5) OPCPackage.save → OutputStream
```

### 5.3 Options（Builder）

| 字段 | 默认 | 含义 |
|------|------|------|
| `strictMode` | false | 严格模式下遇到未映射特性抛异常（保留开关，当前实现尚未在所有不支持特性处触发） |
| `writeSidecar` | true | 关闭后输出纯 xlsx（牺牲 round-trip 保真） |
| `prettyJson` | false | 边车 JSON 是否格式化 |
| `locale` | `EN_US` | 无边车时的 fallback locale |

### 5.4 异常体系

- `UniverXlsxException extends IOException`（checked 顶层）
  - `UniverXlsxReadException`（含 sheet / cell 定位）
  - `UniverXlsxWriteException`
  - `UniverXlsxUnsupportedFeatureException`（严格模式使用）

### 5.5 日志与线程模型

- 使用 `java.util.logging`；级别：`FINE`=字段级 trace，`INFO`=工作簿里程碑，`WARNING`=降级，`SEVERE`=致命。
- `UniverXlsx` 静态方法无共享状态，Reader/Writer 每次调用新建实例；POJO 非线程安全（与普通 POJO 一致）。

## 6. 测试策略

- 框架：JUnit 5 + AssertJ + Jackson（用于 JSON 断言）。
- 三层测试：
  1. 单元测试：工具函数、枚举映射、颜色/长度换算、样式去重 hash。
  2. Converter 集成测试：按字段族拆分（`StyleConverterTest / CellConverterTest / SharedFormulaRegistryTest / RichTextConverterTest / WorksheetConverterTest / SidecarPartTest` 等）。
  3. 端到端 round-trip：`IWorkbookData → 写 xlsx → 读回 → AssertJ recursiveComparison`。

**关键用例**：多 sheet（含 hidden + 乱序 sheetOrder）、cell 类型全覆盖、样式全字段 + 去重、共享公式 si 主格右下、跨 sheet si 不冲突、富文本多 run + 段落、`mergeData` 多形态、冻结三形态 + 无冻结、行列宽高与隐藏、`resources` + `custom` round-trip、未知扩展字段 `extras` 兜底、外部 xlsx（非本库产生）兼容读取且 styles 被填充、空 workbook、sheet id round-trip 不产生 shadow 条目。

**资源**：`src/test/resources/fixtures/*.xlsx`、`src/test/resources/expected/*.json`。

**构建/规范**：
- `./mvnw test` 全绿；`./mvnw -Pcoverage verify` JaCoCo 行覆盖率约 47%（POI 表面较大，主转换路径覆盖良好）。
- `pom.xml` 接入 `p3c-pmd`（`profile=lint`，钉到 `maven-pmd-plugin:3.21.2` 以兼容 PMD 6），`./mvnw -Plint verify` 跑阿里规范检查。
- public 类/方法强制中英双语 Javadoc；测试方法命名 `should_xxx_when_yyy`。

## 7. 目录结构

```
univer-lib/
  pom.xml
  src/main/java/io/github/autoffice/univer/
    UniverXlsx.java
    UniverXlsxOptions.java
    UniverXlsxException.java
    model/ (IWorkbookData.java ...)
    converter/ (WorkbookConverter.java ...)
    io/ (UniverXlsxReader.java, UniverXlsxWriter.java)
    resource/ (SidecarPart.java)
    util/ (ColorUtils.java, LengthUtils.java, IntegerKeyDeserializer.java ...)
  src/test/java/io/github/autoffice/univer/ ...
  example/                # Spring Boot 3 + Vue 3 demo
  docs/design.md          # 本文件
```

## 8. 风险与后续

- POI 依赖体积较大（≈10 MB）；后续若需瘦身可引入「自研 OOXML」作为可插拔 backend。
- 长度换算（px/pt/char）为近似公式，极端情况下 Excel 与 Univer 渲染可能存在 1~2 px 偏差，测试允许容差。
- `pd`、`ol`、`ah`、`scrollTop/Left` 等无 xlsx 原生载体的字段，完全依赖边车；外部 xlsx 读入时会丢失（README 中明示）。
- `strictMode` 当前为预留开关，尚未在所有不支持特性处触发 `UniverXlsxUnsupportedFeatureException`。
