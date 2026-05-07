# univer-lib

Java library for converting between Excel xlsx files and Univer Sheets `IWorkbookData` JSON snapshots.

Univer xlsx 与 `IWorkbookData` 双向高保真转换的 Java 库。

## Features / 特性

- ✅ JDK 8+ 兼容（JDK 8+ compatible）
- ✅ 双向 round-trip：xlsx ↔ `IWorkbookData`
- ✅ 多 sheet、隐藏、`sheetOrder` 排序保留
- ✅ 完整 `IStyleData` 映射：字体、颜色、边框、对齐、换行、旋转、数字格式
- ✅ 共享公式 `si` 主格右下保证
- ✅ 富文本 `IDocumentData`（多 run + 段落）
- ✅ 合并区域、冻结、网格线、RTL、缩放
- ✅ 通过自定义 OPC 边车 `/univer/metadata.json` 无损保留 Univer 专有字段
  (Lossless via custom OPC sidecar `/univer/metadata.json`)

## Maven

```xml
<dependency>
  <groupId>io.github.autoffice</groupId>
  <artifactId>univer-lib</artifactId>
  <version>0.1.0-SNAPSHOT</version>
</dependency>
```

## Quick Start / 快速开始

### Read xlsx → IWorkbookData

```java
import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.IWorkbookData;
import java.nio.file.Paths;

IWorkbookData wb = UniverXlsx.read(Paths.get("input.xlsx"));
Object value = wb.getSheets().get("sheet1").getCellData().get(0).get(0).getV();
```

### Write IWorkbookData → xlsx

```java
import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.*;
import java.nio.file.Paths;
import java.util.*;

IWorkbookData wb = new IWorkbookData()
    .setId("my-workbook")
    .setAppVersion("0.10.2")
    .setLocale("zhCN");

IWorksheetData sheet = new IWorksheetData().setId("s1").setName("Sheet1");
Map<Integer, ICellData> row = new LinkedHashMap<>();
row.put(0, new ICellData().setV("Hello").setT(CellValueType.STRING));
row.put(1, new ICellData().setV(42.0).setT(CellValueType.NUMBER));
sheet.getCellData().put(0, row);

wb.getSheets().put("s1", sheet);
wb.setSheetOrder(Collections.singletonList("s1"));

UniverXlsx.write(wb, Paths.get("output.xlsx"));
```

### Options / 选项

```java
UniverXlsxOptions opts = UniverXlsxOptions.builder()
    .strictMode(false)       // 严格模式（默认 false）
    .writeSidecar(true)      // 是否写入边车（默认 true）
    .prettyJson(false)       // JSON 美化（默认 false）
    .locale("zhCN")         // 缺省 locale
    .build();
UniverXlsx.write(wb, out, opts);
```

## Architecture / 架构

```
IWorkbookData (POJO) ↔ WorkbookConverter ↔ XSSFWorkbook (POI)
                              ↓
  StyleConverter / CellConverter / RichTextConverter
  SharedFormulaRegistry / WorksheetConverter

Univer-specific fields → /univer/metadata.json (OPC sidecar)
```

Univer 专有字段（`resources`、`custom`、`appVersion`、padding、overline 等）xlsx 无原生载体，因此持久化到 OPC 自定义分区 `/univer/metadata.json`，保证读写无损 round-trip。

Univer-specific fields not natively expressible in xlsx are persisted in a custom OPC part for lossless round-trip.

## Limitations / 已知限制

- 长度换算（px ↔ pt ↔ char）为近似值，与 Excel 渲染可能存在 ±1 px 误差
- 极端边框样式（POI BorderStyle 0..13 之外）会被映射到最接近值
- `IStyleData.pd`（padding 上/下/左/右）目前未对齐到 Excel indent，整体落到边车保留
- 公式不计算（仅做字符串 round-trip + cached value）
- 不支持 xlsm（带宏）和 xlsb（二进制）格式
- 严格模式（`strictMode=true`）目前为预留开关，尚未在所有不支持特性处触发 `UniverXlsxUnsupportedFeatureException`

## Build / 构建

```bash
./mvnw test                # 运行所有测试 / run tests
./mvnw -Pcoverage verify   # 生成 JaCoCo 报告 / coverage report at target/site/jacoco/
./mvnw -Plint verify       # 阿里巴巴 p3c 规范检查（见下方说明）
```

### 关于 lint profile / About the lint profile

`-Plint` 接入了阿里巴巴 `p3c-pmd:2.1.1`（基于 PMD 6）。pom 中固定到 `maven-pmd-plugin:3.21.2` 以保证 PMD 6 兼容性。可正常运行。

The `-Plint` profile uses Alibaba `p3c-pmd:2.1.1` (PMD 6 based). The pom pins `maven-pmd-plugin:3.21.2` to ensure PMD 6 compatibility; lint runs successfully.

## Test Coverage / 测试覆盖

当前线覆盖率：约 47%（POI 表面较大，主转换逻辑覆盖良好；POI 内部回退分支与极少触发的 fallback 未覆盖）。

Line coverage: ~47% (POI surface is large; the main conversion paths are well covered; some POI fallback branches and rarely-triggered code paths remain uncovered).

## License

Apache 2.0
