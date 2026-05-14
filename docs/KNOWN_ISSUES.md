# Known Issues — xlsx ↔ Univer IWorkbookData

记录 xlsx ↔ `IWorkbookData` round-trip 过程中**已知会丢失信息**的场景。这些是 Univer 模型本身的表达能力边界（而非实现 bug），除非 Univer 升级模型，否则无法在不破坏 Univer 语义的前提下保留。

修复边界 bug 前，请先在此核对是否为已知限制；修好后请更新本表。

---

## 样式 (Cell Style)

### 1. 非纯色填充图案（FillPattern）丢失

**现象**：
- POI `FillPatternType.SOLID_FOREGROUND` 以外的图案（如 `LEAST_DOTS`、`BRICKS`、`SQUARES`、`FINE_DOTS`、`THICK_HORZ_BANDS` 等 17 种），round-trip 后变为 `NO_FILL`，单元格失去背景图案。

**根因**：
- Univer `IStyleData.bg` 的类型是 `Nullable<IColorStyle>`，仅能表达一个"单色"填充（rgb / theme color）。
- Univer 模型中没有 `fillPattern` 枚举字段，也没有区分 foreground/background 两种颜色的槽位。

**证据**：
- `@univerjs/core` TS 定义：`interface IStyleBase { bg?: Nullable<IColorStyle>; ... }`
- 本项目 `StyleConverter.readBackground`：仅当 `FillPatternType.SOLID_FOREGROUND` 时才写入 `bg`，其他 pattern 主动丢弃。
- 测试固件：`src/test/resources/test2.xlsx` sheet "丰富的单元格表现" 的 `[0,14]`=LEAST_DOTS、`[0,15]`=BRICKS、`[0,16]`=SQUARES round-trip 后 pattern 变为 NO_FILL。

**影响面**：
- Excel 模板里常见的"图案底纹"装饰不会被 Univer UI 展示，重新导出也不会恢复。
- 影响纯展示层；数据、公式、条件格式均不受影响。

**缓解方式（不计划实施）**：
- 可用 sidecar `extras` 保存 `poi_fillPattern = {fillPattern, fg, bg}` 并在写路径还原，从而让 "经过 Univer 再转回 xlsx" 这条链路保真。但 Univer UI 仍不会展示，仅适用于"xlsx → Univer → 打开再另存"这类直通场景。当前策略是"不额外做"。

**覆盖测试**：
- `Test2RoundTripDiffTest` 的 `FILL_PATTERN_NON_SOLID` bucket 已白名单。

---

## 公式 (Formula)

### 2. 非确定公式的 cached value 会漂移

**现象**：
- 含 `RANDBETWEEN`、`RAND`、`TODAY`、`NOW`、`RANDARRAY` 等函数的公式，每次被任何工具（Excel、POI、Univer）重新求值时结果都会变，xlsx 文件里 cached result 与重新求值结果不同属于预期行为。

**根因**：
- 这些函数按 Excel 规范属 *volatile function*，每次 workbook 重新计算时都重新生成值。

**影响面**：
- 读写链路上的 cached value 不可靠；用户肉眼看到的值取决于最后一次打开/保存的时刻。

**处理方式**：
- Round-trip 测试仅比对公式字符串（`A1 语法规范化`），不比对 cached value / evaluated result。
- 应用层若依赖这些函数的结果，应自行触发 `FormulaEvaluator.evaluateAll()` 后读取。

**覆盖测试**：
- `Test2RoundTripDiffTest.readCellValue` 已改为只比对公式字符串。

---

## 其它

后续发现新的模型边界限制时，请追加条目，并说明：现象、根因（引用 Univer TS 类型或文档）、证据、影响面、缓解方式、覆盖测试。
