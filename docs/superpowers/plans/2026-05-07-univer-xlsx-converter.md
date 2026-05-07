# Univer xlsx Converter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 交付 `io.github.autoffice:univer-lib`（Java 8+，Maven），实现 xlsx 与 Univer `IWorkbookData` 双向高保真转换。

**Architecture:** Apache POI (XSSF) 做 OOXML 读写；Univer 特有字段通过 OPC 内部边车分区 `/univer/metadata.json` 可逆保存；Jackson 负责 POJO ↔ JSON；Lombok 简化 POJO。

**Tech Stack:** JDK 8+, Maven, Apache POI 5.x, Jackson Databind 2.x, Lombok, JUnit 5, AssertJ, p3c-pmd（阿里规范）。

**Spec:** `docs/superpowers/specs/2026-05-07-univer-xlsx-converter-design.md`

**Note:** 仓库目前不是 git 仓库。Task 1 会执行 `git init`。所有 commit 步骤假设已 init。

---

## File Structure

```
univer-lib/
  pom.xml
  src/main/java/io/github/autoffice/univer/
    UniverXlsx.java                 — 对外门面
    UniverXlsxOptions.java          — 配置 Builder
    UniverXlsxException.java        — 异常顶层 + 3 个子类
    model/                          — POJO 数据模型（对齐 Univer TS 接口）
      AbstractUniverModel.java      — extras 兜底基类
      BooleanNumber.java            — 0/1 枚举
      CellValueType.java            — 1/2/3/4 枚举
      IColorStyle.java
      ITextDecoration.java
      IBorderStyleData.java
      IBorderData.java
      ITextRotation.java
      IPaddingData.java
      INumfmtLocal.java
      IStyleData.java
      IRange.java
      IRowData.java
      IColumnData.java
      IFreeze.java
      IDocumentData.java
      ICellData.java
      IWorksheetData.java
      IWorkbookData.java
    util/
      IntegerKeyDeserializer.java   — Jackson 数字键反序列化
      ColorUtils.java               — rgb(a) ↔ ARGB hex
      LengthUtils.java              — px / pt / char 近似换算
      JsonMapper.java               — 配置好的 ObjectMapper 单例
    converter/
      StyleConverter.java           — IStyleData ↔ POI CellStyle/Font
      CellConverter.java            — ICellData ↔ POI Cell
      SharedFormulaRegistry.java    — si 分组与主格推导
      RichTextConverter.java        — IDocumentData ↔ XSSFRichTextString
      WorksheetConverter.java       — IWorksheetData ↔ XSSFSheet
      WorkbookConverter.java        — IWorkbookData ↔ XSSFWorkbook
    io/
      UniverXlsxReader.java
      UniverXlsxWriter.java
    resource/
      SidecarPart.java              — OPC 自定义分区读写
  src/test/java/io/github/autoffice/univer/ ...（按组件分 Test 类）
  src/test/resources/
    fixtures/                       — 预生成 xlsx
    expected/                       — 期望 IWorkbookData JSON
```

---

## Task 1: Maven 项目脚手架 + git init

**Files:**
- Create: `pom.xml`
- Create: `.gitignore`

- [ ] **Step 1: 在项目根执行 `git init` 并建立 main 分支**

```bash
cd /Users/aldis/Documents/aldis/code/java/autoffice/univer-lib
git init -b main
```

- [ ] **Step 2: 写入 `.gitignore`**

```
target/
*.class
.idea/
*.iml
.vscode/
.DS_Store
```

- [ ] **Step 3: 写入 `pom.xml`**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd">
  <modelVersion>4.0.0</modelVersion>
  <groupId>io.github.autoffice</groupId>
  <artifactId>univer-lib</artifactId>
  <version>0.1.0-SNAPSHOT</version>
  <packaging>jar</packaging>

  <properties>
    <maven.compiler.source>1.8</maven.compiler.source>
    <maven.compiler.target>1.8</maven.compiler.target>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
    <poi.version>5.2.5</poi.version>
    <jackson.version>2.17.2</jackson.version>
    <lombok.version>1.18.32</lombok.version>
    <junit.version>5.10.2</junit.version>
    <assertj.version>3.25.3</assertj.version>
  </properties>

  <dependencies>
    <dependency><groupId>org.apache.poi</groupId><artifactId>poi</artifactId><version>${poi.version}</version></dependency>
    <dependency><groupId>org.apache.poi</groupId><artifactId>poi-ooxml</artifactId><version>${poi.version}</version></dependency>
    <dependency><groupId>com.fasterxml.jackson.core</groupId><artifactId>jackson-databind</artifactId><version>${jackson.version}</version></dependency>
    <dependency><groupId>org.projectlombok</groupId><artifactId>lombok</artifactId><version>${lombok.version}</version><scope>provided</scope></dependency>
    <dependency><groupId>org.junit.jupiter</groupId><artifactId>junit-jupiter</artifactId><version>${junit.version}</version><scope>test</scope></dependency>
    <dependency><groupId>org.assertj</groupId><artifactId>assertj-core</artifactId><version>${assertj.version}</version><scope>test</scope></dependency>
  </dependencies>

  <build>
    <plugins>
      <plugin><artifactId>maven-surefire-plugin</artifactId><version>3.2.5</version></plugin>
    </plugins>
  </build>

  <profiles>
    <profile>
      <id>coverage</id>
      <build><plugins>
        <plugin>
          <groupId>org.jacoco</groupId><artifactId>jacoco-maven-plugin</artifactId><version>0.8.12</version>
          <executions>
            <execution><goals><goal>prepare-agent</goal></goals></execution>
            <execution><id>report</id><phase>verify</phase><goals><goal>report</goal></goals></execution>
          </executions>
        </plugin>
      </plugins></build>
    </profile>
    <profile>
      <id>lint</id>
      <build><plugins>
        <plugin>
          <groupId>org.apache.maven.plugins</groupId><artifactId>maven-pmd-plugin</artifactId><version>3.22.0</version>
          <configuration><rulesets><ruleset>rulesets/java/ali-comment.xml</ruleset><ruleset>rulesets/java/ali-naming.xml</ruleset></rulesets></configuration>
          <dependencies><dependency><groupId>com.alibaba.p3c</groupId><artifactId>p3c-pmd</artifactId><version>2.1.1</version></dependency></dependencies>
          <executions><execution><phase>verify</phase><goals><goal>check</goal></goals></execution></executions>
        </plugin>
      </plugins></build>
    </profile>
  </profiles>
</project>
```

- [ ] **Step 4: 验证构建**

Run: `mvn -q -DskipTests package`
Expected: `BUILD SUCCESS`，生成 `target/univer-lib-0.1.0-SNAPSHOT.jar`（空 jar 也算成功）。

- [ ] **Step 5: Commit**

```bash
git add pom.xml .gitignore
git commit -m "chore: scaffold maven project for univer-lib"
```

---

## Task 2: 枚举与基础 POJO 基类

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/model/BooleanNumber.java`
- Create: `src/main/java/io/github/autoffice/univer/model/CellValueType.java`
- Create: `src/main/java/io/github/autoffice/univer/model/AbstractUniverModel.java`
- Create: `src/main/java/io/github/autoffice/univer/util/JsonMapper.java`
- Test: `src/test/java/io/github/autoffice/univer/model/BooleanNumberTest.java`
- Test: `src/test/java/io/github/autoffice/univer/model/AbstractUniverModelTest.java`

- [ ] **Step 1: 写失败测试 `BooleanNumberTest`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class BooleanNumberTest {
    @Test
    void should_serialize_as_number_and_deserialize_back() throws Exception {
        ObjectMapper m = JsonMapper.get();
        assertThat(m.writeValueAsString(BooleanNumber.TRUE)).isEqualTo("1");
        assertThat(m.writeValueAsString(BooleanNumber.FALSE)).isEqualTo("0");
        assertThat(m.readValue("1", BooleanNumber.class)).isEqualTo(BooleanNumber.TRUE);
        assertThat(m.readValue("0", BooleanNumber.class)).isEqualTo(BooleanNumber.FALSE);
    }
}
```

- [ ] **Step 2: 运行测试确认失败**

Run: `mvn -q test -Dtest=BooleanNumberTest`
Expected: 编译失败（类不存在）。

- [ ] **Step 3: 实现 `BooleanNumber`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

/** Univer 中用数字 0/1 表达布尔语义的枚举 / boolean-as-number enum used by Univer model. */
public enum BooleanNumber {
    /** 假 / false. */
    FALSE(0),
    /** 真 / true. */
    TRUE(1);

    private final int value;

    BooleanNumber(int value) { this.value = value; }

    @JsonValue
    public int getValue() { return value; }

    @JsonCreator
    public static BooleanNumber of(int v) { return v == 0 ? FALSE : TRUE; }
}
```

- [ ] **Step 4: 实现 `CellValueType`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

/** 单元格值类型 / cell value type (Univer CellValueType). */
public enum CellValueType {
    STRING(1), NUMBER(2), BOOLEAN(3), FORCE_TEXT(4);

    private final int value;
    CellValueType(int v) { this.value = v; }
    @JsonValue public int getValue() { return value; }
    @JsonCreator public static CellValueType of(int v) {
        for (CellValueType t : values()) { if (t.value == v) return t; }
        throw new IllegalArgumentException("Unknown CellValueType: " + v);
    }
}
```

- [ ] **Step 5: 实现 `JsonMapper`**

```java
package io.github.autoffice.univer.util;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

/** 库内统一的 Jackson ObjectMapper / shared ObjectMapper for the library. */
public final class JsonMapper {
    private static final ObjectMapper MAPPER = new ObjectMapper()
            .setSerializationInclusion(JsonInclude.Include.NON_NULL)
            .disable(SerializationFeature.FAIL_ON_EMPTY_BEANS)
            .disable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES);

    private JsonMapper() {}

    public static ObjectMapper get() { return MAPPER; }
}
```

- [ ] **Step 6: 实现 `AbstractUniverModel`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonAnyGetter;
import com.fasterxml.jackson.annotation.JsonAnySetter;
import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/** 所有 Univer POJO 的基类，承接未知字段以防版本升级丢数据 /
 *  Base class that captures unknown fields via extras map for forward compatibility. */
public abstract class AbstractUniverModel {
    @JsonIgnore
    private final Map<String, Object> extras = new LinkedHashMap<>();

    @JsonAnyGetter
    public Map<String, Object> getExtras() { return extras; }

    @JsonAnySetter
    public void putExtra(String key, Object value) { extras.put(key, value); }

    @Override public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        return Objects.equals(extras, ((AbstractUniverModel) o).extras);
    }
    @Override public int hashCode() { return Objects.hash(extras); }
}
```

- [ ] **Step 7: 写 `AbstractUniverModelTest`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;
import io.github.autoffice.univer.util.JsonMapper;

class AbstractUniverModelTest {
    @Data @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    static class Sample extends AbstractUniverModel { private String name; }

    @Test
    void should_roundtrip_unknown_fields_via_extras() throws Exception {
        Sample s = JsonMapper.get().readValue("{\"name\":\"a\",\"future\":42}", Sample.class);
        assertThat(s.getName()).isEqualTo("a");
        assertThat(s.getExtras()).containsEntry("future", 42);
        String json = JsonMapper.get().writeValueAsString(s);
        assertThat(json).contains("\"future\":42");
    }
}
```

- [ ] **Step 8: 运行全部测试**

Run: `mvn -q test`
Expected: 全部测试通过。

- [ ] **Step 9: Commit**

```bash
git add src
git commit -m "feat(model): add BooleanNumber, CellValueType, AbstractUniverModel, JsonMapper"
```

---

## Task 3: 叶子 POJO（颜色/装饰/边框/旋转/内边距/数字格式/区间/行列/冻结）

**Files（全部 Create）:**
- `src/main/java/io/github/autoffice/univer/model/IColorStyle.java`
- `src/main/java/io/github/autoffice/univer/model/ITextDecoration.java`
- `src/main/java/io/github/autoffice/univer/model/IBorderStyleData.java`
- `src/main/java/io/github/autoffice/univer/model/IBorderData.java`
- `src/main/java/io/github/autoffice/univer/model/ITextRotation.java`
- `src/main/java/io/github/autoffice/univer/model/IPaddingData.java`
- `src/main/java/io/github/autoffice/univer/model/INumfmtLocal.java`
- `src/main/java/io/github/autoffice/univer/model/IRange.java`
- `src/main/java/io/github/autoffice/univer/model/IRowData.java`
- `src/main/java/io/github/autoffice/univer/model/IColumnData.java`
- `src/main/java/io/github/autoffice/univer/model/IFreeze.java`
- Test: `src/test/java/io/github/autoffice/univer/model/LeafPojoJsonTest.java`

每个 POJO 统一写法（示例 `IColorStyle`）：

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/** 颜色样式 / color style. */
@Data
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IColorStyle extends AbstractUniverModel {
    /** RGB 颜色（#rrggbb） / rgb color. */
    private String rgb;
    /** 主题色索引 / theme color index. */
    private Integer th;
}
```

- [ ] **Step 1: 写失败测试 `LeafPojoJsonTest`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class LeafPojoJsonTest {
    private final ObjectMapper m = JsonMapper.get();

    @Test
    void should_roundtrip_color_style() throws Exception {
        IColorStyle c = new IColorStyle().setRgb("#ff0000");
        String json = m.writeValueAsString(c);
        assertThat(json).isEqualTo("{\"rgb\":\"#ff0000\"}");
        assertThat(m.readValue(json, IColorStyle.class)).isEqualTo(c);
    }

    @Test
    void should_roundtrip_border_data() throws Exception {
        IBorderData bd = new IBorderData().setT(new IBorderStyleData().setS(1)
            .setCl(new IColorStyle().setRgb("#000000")));
        String json = m.writeValueAsString(bd);
        assertThat(m.readValue(json, IBorderData.class)).isEqualTo(bd);
    }

    @Test
    void should_roundtrip_range() throws Exception {
        IRange r = new IRange().setStartRow(0).setStartColumn(0).setEndRow(2).setEndColumn(3);
        String json = m.writeValueAsString(r);
        assertThat(m.readValue(json, IRange.class)).isEqualTo(r);
    }
}
```

- [ ] **Step 2: 运行确认失败**

Run: `mvn -q test -Dtest=LeafPojoJsonTest`
Expected: 编译失败。

- [ ] **Step 3: 实现 `ITextDecoration`（ul/st/ol 复用）**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data; import lombok.EqualsAndHashCode; import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/** 文字装饰（下划线 / 删除线 / 上划线） / text decoration (underline/strike/overline). */
@Data @NoArgsConstructor @Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ITextDecoration extends AbstractUniverModel {
    private BooleanNumber s;
    private BooleanNumber c;
    private IColorStyle cl;
    private Integer t;
}
```

- [ ] **Step 4: 实现 `IBorderStyleData`**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data; import lombok.EqualsAndHashCode; import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/** 单边边框样式 / single-side border style. */
@Data @NoArgsConstructor @Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IBorderStyleData extends AbstractUniverModel {
    /** 边框类型索引 / border style index. */
    private Integer s;
    /** 边框颜色 / border color. */
    private IColorStyle cl;
}
```

- [ ] **Step 5: 实现 `IBorderData`（t/b/l/r/tl/tr/bl/br）**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data; import lombok.EqualsAndHashCode; import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/** 边框集合（上/下/左/右/四角对角线） / border set. */
@Data @NoArgsConstructor @Accessors(chain = true)
@EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IBorderData extends AbstractUniverModel {
    private IBorderStyleData t;
    private IBorderStyleData b;
    private IBorderStyleData l;
    private IBorderStyleData r;
    private IBorderStyleData tl;
    private IBorderStyleData tr;
    private IBorderStyleData bl;
    private IBorderStyleData br;
}
```

- [ ] **Step 6: 实现 `ITextRotation`、`IPaddingData`、`INumfmtLocal`**

```java
// ITextRotation.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ITextRotation extends AbstractUniverModel {
    private Integer a;           // 旋转角度 / rotation angle
    private BooleanNumber v;     // 是否竖排 / vertical
}
```

```java
// IPaddingData.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IPaddingData extends AbstractUniverModel {
    private Integer t; private Integer b; private Integer l; private Integer r;
}
```

```java
// INumfmtLocal.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class INumfmtLocal extends AbstractUniverModel {
    /** 数字格式 / number format pattern. */ private String pattern;
}
```

- [ ] **Step 7: 实现 `IRange`、`IRowData`、`IColumnData`、`IFreeze`**

```java
// IRange.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IRange extends AbstractUniverModel {
    private Integer startRow; private Integer startColumn;
    private Integer endRow;   private Integer endColumn;
    private Integer rangeType; private String sheetId;
}
```

```java
// IRowData.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IRowData extends AbstractUniverModel {
    private Double h;                   // 行高（px） / row height in px
    private Double ah;                  // 自动行高 / auto row height
    private BooleanNumber hd;           // 是否隐藏 / hidden
}
```

```java
// IColumnData.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IColumnData extends AbstractUniverModel {
    private Double w; private BooleanNumber hd;
}
```

```java
// IFreeze.java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IFreeze extends AbstractUniverModel {
    private Integer startRow; private Integer startColumn;
    private Integer xSplit;   private Integer ySplit;
}
```

- [ ] **Step 8: 运行测试**

Run: `mvn -q test -Dtest=LeafPojoJsonTest`
Expected: PASS。

- [ ] **Step 9: Commit**

```bash
git add src
git commit -m "feat(model): add leaf POJOs (color/border/rotation/padding/numfmt/range/row/column/freeze)"
```

---

## Task 4: 复合 POJO（IStyleData / IDocumentData / ICellData / IWorksheetData / IWorkbookData）

**Files（全部 Create）:**
- `src/main/java/io/github/autoffice/univer/model/IStyleData.java`
- `src/main/java/io/github/autoffice/univer/model/IDocumentData.java`
- `src/main/java/io/github/autoffice/univer/model/ICellData.java`
- `src/main/java/io/github/autoffice/univer/model/IWorksheetData.java`
- `src/main/java/io/github/autoffice/univer/model/IWorkbookData.java`
- `src/main/java/io/github/autoffice/univer/util/IntegerKeyDeserializer.java`
- Test: `src/test/java/io/github/autoffice/univer/model/CompositePojoJsonTest.java`

- [ ] **Step 1: 实现 `IntegerKeyDeserializer`**

```java
package io.github.autoffice.univer.util;

import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.KeyDeserializer;

/** 把 JSON 对象中的数字字符串 key 反序列化为 Integer / deserialize numeric string keys to Integer. */
public class IntegerKeyDeserializer extends KeyDeserializer {
    @Override
    public Object deserializeKey(String key, DeserializationContext ctxt) {
        return Integer.parseInt(key);
    }
}
```

- [ ] **Step 2: 实现 `IStyleData`**

```java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;

/** 单元格样式 / cell style. */
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IStyleData extends AbstractUniverModel {
    private String ff;              // 字体 / font family
    private Integer fs;             // 字号 / font size (pt)
    private BooleanNumber it;       // 斜体 / italic
    private BooleanNumber bl;       // 粗体 / bold
    private ITextDecoration ul;     // 下划线 / underline
    private ITextDecoration st;     // 删除线 / strikethrough
    private ITextDecoration ol;     // 上划线 / overline
    private IColorStyle bg;         // 背景色 / background
    private IBorderData bd;         // 边框 / border
    private IColorStyle cl;         // 字体色 / font color
    private Integer va;             // 上下标 / super/sub script
    private ITextRotation tr;       // 旋转 / rotation
    private Integer ht;             // 水平对齐 / horizontal align
    private Integer vt;             // 垂直对齐 / vertical align
    private Integer tb;             // 溢出策略 / wrap strategy
    private IPaddingData pd;        // 内边距 / padding
    private INumfmtLocal n;         // 数字格式 / number format
}
```

- [ ] **Step 3: 实现 `IDocumentData`（富文本最小可 round-trip 结构）**

```java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
import java.util.List;
import java.util.Map;

/** 富文本文档 / rich text document (Univer Doc minimal shape). */
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IDocumentData extends AbstractUniverModel {
    private Body body;
    private String id;
    private Map<String, Object> documentStyle;

    /** 富文本正文 / rich text body. */
    @Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Body extends AbstractUniverModel {
        private String dataStream;
        private List<TextRun> textRuns;
        private List<Paragraph> paragraphs;
    }
    /** 文本 run / text run with style range. */
    @Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class TextRun extends AbstractUniverModel {
        private Integer st; private Integer ed; private IStyleData ts;
    }
    /** 段落 / paragraph break. */
    @Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Paragraph extends AbstractUniverModel {
        private Integer startIndex;
    }
}
```

- [ ] **Step 4: 实现 `ICellData`**

```java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
import java.util.Map;

/** 单元格数据 / cell data. */
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ICellData extends AbstractUniverModel {
    /** 原始值（字符串/数字/布尔） / raw value (string/number/boolean). */
    private Object v;
    /** 样式 id 或样式对象 / style id (string) or inline style object. */
    private Object s;
    /** 类型 / cell value type. */
    private CellValueType t;
    /** 富文本 / rich text. */
    private IDocumentData p;
    /** 公式 / formula expression. */
    private String f;
    /** 公式 id / formula id for shared formulas. */
    private String si;
    /** 自定义字段 / custom payload. */
    private Map<String, Object> custom;
}
```

- [ ] **Step 5: 实现 `IWorksheetData`**

```java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import io.github.autoffice.univer.util.IntegerKeyDeserializer;
import lombok.*; import lombok.experimental.Accessors;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** 工作表数据 / worksheet data. */
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IWorksheetData extends AbstractUniverModel {
    private String id;
    private String name;
    private String tabColor;
    private BooleanNumber hidden;
    private IFreeze freeze;
    private Integer rowCount;
    private Integer columnCount;
    private Double defaultColumnWidth;
    private Double defaultRowHeight;
    private List<IRange> mergeData;

    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class,
                     contentUsing = com.fasterxml.jackson.databind.deser.std.MapDeserializer.class)
    private Map<Integer, Map<Integer, ICellData>> cellData = new LinkedHashMap<>();

    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class)
    private Map<Integer, IRowData> rowData = new LinkedHashMap<>();

    @JsonDeserialize(keyUsing = IntegerKeyDeserializer.class)
    private Map<Integer, IColumnData> columnData = new LinkedHashMap<>();

    private Header rowHeader;
    private Header columnHeader;
    private BooleanNumber showGridlines;
    private BooleanNumber rightToLeft;
    private IStyleData defaultStyle;
    private Double zoomRatio;
    private Double scrollTop;
    private Double scrollLeft;

    /** 行/列头配置 / row/column header config. */
    @Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    public static class Header extends AbstractUniverModel {
        private Double width; private Double height; private BooleanNumber hidden;
    }
}
```

备注：`cellData` 的嵌套 Map 反序列化在 Jackson 中默认即可处理（内层 key 仍是数字字符串 → 通过同一 `IntegerKeyDeserializer` 生效需在 ObjectMapper 层注册）。在 `JsonMapper` 中追加：`.setDefaultKeyDeserializer(new IntegerKeyDeserializer())` 仅当该测试失败时再加，或改为显式在 `@JsonDeserialize` 注解上配置。实际实现按「默认全局 KeyDeserializer 为 IntegerKeyDeserializer」方案在 Step 8 调整 `JsonMapper` 并补测试。

- [ ] **Step 6: 实现 `IWorkbookData`**

```java
package io.github.autoffice.univer.model;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.*; import lombok.experimental.Accessors;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** 工作簿数据 / workbook data (root snapshot). */
@Data @NoArgsConstructor @Accessors(chain = true) @EqualsAndHashCode(callSuper = true)
@JsonInclude(JsonInclude.Include.NON_NULL)
public class IWorkbookData extends AbstractUniverModel {
    private String id;
    private String name;
    private String appVersion;
    private String locale;
    private Map<String, IStyleData> styles = new LinkedHashMap<>();
    private List<String> sheetOrder;
    private Map<String, IWorksheetData> sheets = new LinkedHashMap<>();
    private Object resources;
}
```

- [ ] **Step 7: 写 `CompositePojoJsonTest`（含 cellData 嵌套 key 反序列化）**

```java
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class CompositePojoJsonTest {
    private final ObjectMapper m = JsonMapper.get();

    @Test
    void should_deserialize_cell_data_with_numeric_keys() throws Exception {
        String json = "{\"cellData\":{\"0\":{\"0\":{\"v\":\"A1\"},\"1\":{\"v\":1}}}}";
        IWorksheetData ws = m.readValue(json, IWorksheetData.class);
        assertThat(ws.getCellData().get(0).get(0).getV()).isEqualTo("A1");
        assertThat(ws.getCellData().get(0).get(1).getV()).isEqualTo(1);
    }

    @Test
    void should_roundtrip_minimal_workbook() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb1").setName("demo")
            .setAppVersion("0.10.2").setLocale("enUS");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        wb.setSheetOrder(java.util.Collections.singletonList("s1"));
        String json = m.writeValueAsString(wb);
        IWorkbookData back = m.readValue(json, IWorkbookData.class);
        assertThat(back).isEqualTo(wb);
    }
}
```

- [ ] **Step 8: 在 `JsonMapper` 中注册全局 `IntegerKeyDeserializer`**

编辑 `JsonMapper.java`，追加：

```java
import com.fasterxml.jackson.databind.module.SimpleModule;
import io.github.autoffice.univer.util.IntegerKeyDeserializer;

// 在 MAPPER 初始化链式调用后追加：
static {
    SimpleModule mod = new SimpleModule();
    mod.addKeyDeserializer(Integer.class, new IntegerKeyDeserializer());
    MAPPER.registerModule(mod);
}
```

- [ ] **Step 9: 运行测试**

Run: `mvn -q test`
Expected: 全绿。

- [ ] **Step 10: Commit**

```bash
git add src
git commit -m "feat(model): add composite POJOs (style/doc/cell/worksheet/workbook)"
```

---

## Task 5: 工具类（颜色 + 长度换算）

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/util/ColorUtils.java`
- Create: `src/main/java/io/github/autoffice/univer/util/LengthUtils.java`
- Test: `src/test/java/io/github/autoffice/univer/util/ColorUtilsTest.java`
- Test: `src/test/java/io/github/autoffice/univer/util/LengthUtilsTest.java`

- [ ] **Step 1: 写失败测试 `ColorUtilsTest`**

```java
package io.github.autoffice.univer.util;

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class ColorUtilsTest {
    @Test
    void should_convert_rgb_to_argb_bytes() {
        byte[] argb = ColorUtils.rgbHexToArgb("#ff0000");
        assertThat(argb).containsExactly((byte) 0xFF, (byte) 0xFF, 0, 0);
    }

    @Test
    void should_convert_argb_to_rgb_hex() {
        assertThat(ColorUtils.argbToRgbHex(new byte[]{(byte)0xFF,(byte)0xFF,0,0}))
            .isEqualTo("#ff0000");
    }

    @Test
    void should_strip_alpha_for_rgba() {
        byte[] argb = ColorUtils.rgbHexToArgb("#80ff0000");
        assertThat(argb[0]).isEqualTo((byte) 0x80);
    }
}
```

- [ ] **Step 2: 运行确认失败**

Run: `mvn -q test -Dtest=ColorUtilsTest`
Expected: 编译失败。

- [ ] **Step 3: 实现 `ColorUtils`**

```java
package io.github.autoffice.univer.util;

/** 颜色转换工具 / color conversion utilities. */
public final class ColorUtils {
    private ColorUtils() {}

    /** `#rrggbb` 或 `#aarrggbb` → ARGB 字节数组 / hex to ARGB bytes. */
    public static byte[] rgbHexToArgb(String hex) {
        if (hex == null) return null;
        String s = hex.startsWith("#") ? hex.substring(1) : hex;
        int a = 0xFF, r, g, b;
        if (s.length() == 6) {
            r = Integer.parseInt(s.substring(0, 2), 16);
            g = Integer.parseInt(s.substring(2, 4), 16);
            b = Integer.parseInt(s.substring(4, 6), 16);
        } else if (s.length() == 8) {
            a = Integer.parseInt(s.substring(0, 2), 16);
            r = Integer.parseInt(s.substring(2, 4), 16);
            g = Integer.parseInt(s.substring(4, 6), 16);
            b = Integer.parseInt(s.substring(6, 8), 16);
        } else {
            throw new IllegalArgumentException("Invalid hex color: " + hex);
        }
        return new byte[]{(byte) a, (byte) r, (byte) g, (byte) b};
    }

    /** ARGB 字节 → `#rrggbb`（透明度丢弃） / ARGB bytes to `#rrggbb`. */
    public static String argbToRgbHex(byte[] argb) {
        if (argb == null || argb.length < 4) return null;
        return String.format("#%02x%02x%02x", argb[1] & 0xFF, argb[2] & 0xFF, argb[3] & 0xFF);
    }
}
```

- [ ] **Step 4: 写失败测试 `LengthUtilsTest`**

```java
package io.github.autoffice.univer.util;

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;

class LengthUtilsTest {
    @Test
    void should_convert_px_to_points() {
        assertThat(LengthUtils.pxToPoints(96.0)).isCloseTo(72.0, within(0.01));
    }
    @Test
    void should_convert_points_to_px() {
        assertThat(LengthUtils.pointsToPx(72.0)).isCloseTo(96.0, within(0.01));
    }
    @Test
    void should_convert_px_to_chars() {
        assertThat(LengthUtils.pxToChars(70)).isCloseTo(10.0, within(0.5));
    }
}
```

- [ ] **Step 5: 实现 `LengthUtils`**

```java
package io.github.autoffice.univer.util;

/** 长度换算（近似） / length conversion utilities (approximate). */
public final class LengthUtils {
    private static final double PX_PER_INCH = 96.0;
    private static final double PT_PER_INCH = 72.0;
    private static final double PX_PER_CHAR = 7.0;

    private LengthUtils() {}

    /** px → pt. */
    public static double pxToPoints(double px) { return px * PT_PER_INCH / PX_PER_INCH; }
    /** pt → px. */
    public static double pointsToPx(double pt) { return pt * PX_PER_INCH / PT_PER_INCH; }
    /** px → 字符数（Excel 列宽单位） / px to chars (Excel column-width unit). */
    public static double pxToChars(double px) { return px / PX_PER_CHAR; }
    /** 字符数 → px / chars to px. */
    public static double charsToPx(double chars) { return chars * PX_PER_CHAR; }
}
```

- [ ] **Step 6: 运行测试**

Run: `mvn -q test`
Expected: 全绿。

- [ ] **Step 7: Commit**

```bash
git add src
git commit -m "feat(util): add ColorUtils and LengthUtils"
```

---

## Task 6: 对外 API 外壳与异常体系

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsxException.java`
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsxReadException.java`
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsxWriteException.java`
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsxUnsupportedFeatureException.java`
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsxOptions.java`
- Create: `src/main/java/io/github/autoffice/univer/UniverXlsx.java`
- Test: `src/test/java/io/github/autoffice/univer/UniverXlsxApiTest.java`

- [ ] **Step 1: 写异常类（4 个）**

```java
package io.github.autoffice.univer;
import java.io.IOException;

/** 库顶层异常 / top-level exception. */
public class UniverXlsxException extends IOException {
    public UniverXlsxException(String msg) { super(msg); }
    public UniverXlsxException(String msg, Throwable cause) { super(msg, cause); }
}
```

```java
package io.github.autoffice.univer;
/** 读取异常 / read-side exception. */
public class UniverXlsxReadException extends UniverXlsxException {
    public UniverXlsxReadException(String msg) { super(msg); }
    public UniverXlsxReadException(String msg, Throwable cause) { super(msg, cause); }
}
```

```java
package io.github.autoffice.univer;
/** 写出异常 / write-side exception. */
public class UniverXlsxWriteException extends UniverXlsxException {
    public UniverXlsxWriteException(String msg) { super(msg); }
    public UniverXlsxWriteException(String msg, Throwable cause) { super(msg, cause); }
}
```

```java
package io.github.autoffice.univer;
/** 严格模式下遇到不支持特性的异常 / unsupported-feature exception in strict mode. */
public class UniverXlsxUnsupportedFeatureException extends UniverXlsxException {
    public UniverXlsxUnsupportedFeatureException(String msg) { super(msg); }
}
```

- [ ] **Step 2: 实现 `UniverXlsxOptions`（Builder）**

```java
package io.github.autoffice.univer;

import lombok.Builder;
import lombok.Getter;

/** 读写选项 / reader/writer options. */
@Getter
@Builder(toBuilder = true)
public class UniverXlsxOptions {
    /** 严格模式 / strict mode (throw on unsupported features). */
    @Builder.Default private boolean strictMode = false;
    /** 是否写入边车 / write sidecar part. */
    @Builder.Default private boolean writeSidecar = true;
    /** 美化 JSON / pretty-print sidecar json. */
    @Builder.Default private boolean prettyJson = false;
    /** 缺省 locale / fallback locale when sidecar missing. */
    @Builder.Default private String locale = "enUS";

    /** 默认配置 / default options. */
    public static UniverXlsxOptions defaults() { return UniverXlsxOptions.builder().build(); }
}
```

- [ ] **Step 3: 实现 `UniverXlsx` 外壳（方法体先抛 UnsupportedOperationException）**

```java
package io.github.autoffice.univer;

import io.github.autoffice.univer.io.UniverXlsxReader;
import io.github.autoffice.univer.io.UniverXlsxWriter;
import io.github.autoffice.univer.model.IWorkbookData;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;

/** Univer xlsx 转换门面 / facade for xlsx ↔ IWorkbookData conversion. */
public final class UniverXlsx {
    private UniverXlsx() {}

    public static IWorkbookData read(InputStream in) throws UniverXlsxException {
        return read(in, UniverXlsxOptions.defaults());
    }
    public static IWorkbookData read(Path path) throws UniverXlsxException {
        try (InputStream in = Files.newInputStream(path)) {
            return read(in, UniverXlsxOptions.defaults());
        } catch (IOException e) { throw new UniverXlsxReadException("open file failed: " + path, e); }
    }
    public static IWorkbookData read(InputStream in, UniverXlsxOptions opts) throws UniverXlsxException {
        return new UniverXlsxReader(opts).read(in);
    }

    public static void write(IWorkbookData wb, OutputStream out) throws UniverXlsxException {
        write(wb, out, UniverXlsxOptions.defaults());
    }
    public static void write(IWorkbookData wb, Path path) throws UniverXlsxException {
        try (OutputStream out = Files.newOutputStream(path)) {
            write(wb, out, UniverXlsxOptions.defaults());
        } catch (IOException e) { throw new UniverXlsxWriteException("open file failed: " + path, e); }
    }
    public static void write(IWorkbookData wb, OutputStream out, UniverXlsxOptions opts) throws UniverXlsxException {
        new UniverXlsxWriter(opts).write(wb, out);
    }
}
```

- [ ] **Step 4: 建立 `UniverXlsxReader / UniverXlsxWriter` 空壳使编译通过**

```java
// src/main/java/io/github/autoffice/univer/io/UniverXlsxReader.java
package io.github.autoffice.univer.io;
import io.github.autoffice.univer.*;
import io.github.autoffice.univer.model.IWorkbookData;
import java.io.InputStream;
public class UniverXlsxReader {
    private final UniverXlsxOptions opts;
    public UniverXlsxReader(UniverXlsxOptions opts) { this.opts = opts; }
    public IWorkbookData read(InputStream in) throws UniverXlsxException {
        throw new UniverXlsxReadException("not implemented");
    }
}
```

```java
// src/main/java/io/github/autoffice/univer/io/UniverXlsxWriter.java
package io.github.autoffice.univer.io;
import io.github.autoffice.univer.*;
import io.github.autoffice.univer.model.IWorkbookData;
import java.io.OutputStream;
public class UniverXlsxWriter {
    private final UniverXlsxOptions opts;
    public UniverXlsxWriter(UniverXlsxOptions opts) { this.opts = opts; }
    public void write(IWorkbookData wb, OutputStream out) throws UniverXlsxException {
        throw new UniverXlsxWriteException("not implemented");
    }
}
```

- [ ] **Step 5: 写 `UniverXlsxApiTest`**

```java
package io.github.autoffice.univer;

import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class UniverXlsxApiTest {
    @Test
    void should_expose_read_and_write_facade() {
        assertThatThrownBy(() -> UniverXlsx.read(new ByteArrayInputStream(new byte[0])))
            .isInstanceOf(UniverXlsxReadException.class);
        assertThatThrownBy(() -> UniverXlsx.write(new io.github.autoffice.univer.model.IWorkbookData(),
                                                   new ByteArrayOutputStream()))
            .isInstanceOf(UniverXlsxWriteException.class);
    }
}
```

- [ ] **Step 6: 运行测试**

Run: `mvn -q test`
Expected: 全绿。

- [ ] **Step 7: Commit**

```bash
git add src
git commit -m "feat(api): add facade, options and exception hierarchy"
```

---

## Task 7: 样式转换器 StyleConverter

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/StyleConverter.java`
- Test: `src/test/java/io/github/autoffice/univer/converter/StyleConverterTest.java`

**接口契约**（实现时完全按此签名，其它任务依赖）：

```java
public final class StyleConverter {
    public StyleConverter(XSSFWorkbook wb);                         // ctor，内部缓存 style hash → XSSFCellStyle
    public XSSFCellStyle toPoiStyle(IStyleData s);                   // 写：IStyleData → POI CellStyle（带缓存）
    public IStyleData fromPoiStyle(XSSFCellStyle cs);                // 读：POI CellStyle → IStyleData
    public String styleIdOf(IStyleData s);                           // 读路径的样式去重：对 IStyleData 生成稳定 id
}
```

- [ ] **Step 1: 写测试 `StyleConverterTest`（覆盖字体、颜色、对齐、边框、换行、旋转、数字格式）**

```java
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class StyleConverterTest {

    @Test
    void should_roundtrip_font_and_alignment() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setFf("Arial").setFs(12)
                .setBl(BooleanNumber.TRUE).setIt(BooleanNumber.TRUE)
                .setHt(1).setVt(2).setCl(new IColorStyle().setRgb("#ff0000"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getAlignment()).isEqualTo(HorizontalAlignment.LEFT);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getFf()).isEqualTo("Arial");
            assertThat(back.getFs()).isEqualTo(12);
            assertThat(back.getBl()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getIt()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getCl().getRgb()).isEqualToIgnoringCase("#ff0000");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_roundtrip_background_border_wrap_rotation() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData()
                .setBg(new IColorStyle().setRgb("#00ff00"))
                .setBd(new IBorderData().setT(new IBorderStyleData().setS(1)
                        .setCl(new IColorStyle().setRgb("#000000"))))
                .setTb(3)
                .setTr(new ITextRotation().setA(45));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getBg().getRgb()).isEqualToIgnoringCase("#00ff00");
            assertThat(back.getBd().getT().getS()).isEqualTo(1);
            assertThat(back.getTb()).isEqualTo(3);
            assertThat(back.getTr().getA()).isEqualTo(45);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_roundtrip_number_format() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setN(new INumfmtLocal().setPattern("yyyy-mm-dd"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getDataFormatString()).isEqualTo("yyyy-mm-dd");
            assertThat(sc.fromPoiStyle(poi).getN().getPattern()).isEqualTo("yyyy-mm-dd");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_cache_equal_styles_to_one_poi_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData a = new IStyleData().setFf("Arial").setFs(10);
            IStyleData b = new IStyleData().setFf("Arial").setFs(10);
            assertThat(sc.toPoiStyle(a)).isSameAs(sc.toPoiStyle(b));
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_generate_stable_style_id() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData a = new IStyleData().setFf("Arial").setFs(10);
            IStyleData b = new IStyleData().setFf("Arial").setFs(10);
            assertThat(sc.styleIdOf(a)).isEqualTo(sc.styleIdOf(b));
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
```

- [ ] **Step 2: 运行确认失败**

Run: `mvn -q test -Dtest=StyleConverterTest`
Expected: 编译失败。

- [ ] **Step 3: 实现 `StyleConverter`**

代码较长，放在 `src/main/java/io/github/autoffice/univer/converter/StyleConverter.java`。关键要点：

- 内部 `Map<IStyleData, XSSFCellStyle> cache` 做写侧缓存；`styleIdOf` 用 SHA-1(json(IStyleData)) 截前 16 hex 字符。
- 字体：`toPoiStyle` 新建 `XSSFFont`，按 `ff/fs/bl/it/ul(.s=1→SINGLE)/st(.s=1→true)/cl.rgb/va` 设置。
- 颜色：写侧 `new XSSFColor(ColorUtils.rgbHexToArgb(hex), null)`，读侧取 `XSSFColor.getARGB()` 再 `argbToRgbHex`。
- 背景：`setFillForegroundColor(new XSSFColor(...))` + `setFillPattern(FillPatternType.SOLID_FOREGROUND)`。
- 边框：8 个方向按 `s` 索引映射 `BorderStyle.values()[s]`；`cl` 同颜色处理。
- `ht/vt`：`HorizontalAlignment` 1=LEFT, 2=CENTER, 3=RIGHT；`VerticalAlignment` 1=TOP, 2=CENTER, 3=BOTTOM。
- `tb`：1→`setWrapText(false)`；2→`setShrinkToFit(true)+setWrapText(false)`；3→`setWrapText(true)`。
- 旋转：`tr.v=1` → rotation=255；否则 `rotation=min(a,180)`。
- 数字格式：`wb.createDataFormat().getFormat(pattern)` 后 `cellStyle.setDataFormat(fmt)`。
- 下划线的 `t` 字段（linethrough 类型）xlsx 原生仅支持 SINGLE/DOUBLE，映射关系按 0..4 简单对齐；不能精确还原的值写入 `extras`。

实现完成后运行 `mvn -q test -Dtest=StyleConverterTest` 直至全部通过。

- [ ] **Step 4: Commit**

```bash
git add src
git commit -m "feat(converter): add StyleConverter for IStyleData <-> XSSFCellStyle"
```

---

## Task 8: 单元格转换器 CellConverter（含类型推断、强文本、自定义字段）

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/CellConverter.java`
- Test: `src/test/java/io/github/autoffice/univer/converter/CellConverterTest.java`

**接口**：

```java
public final class CellConverter {
    public CellConverter(StyleConverter styles);
    public void writeCell(XSSFCell poiCell, ICellData src);                  // ICellData → POI Cell
    public ICellData readCell(XSSFCell poiCell);                              // POI Cell → ICellData（不含 si 处理，si 由 SharedFormulaRegistry 负责）
}
```

- [ ] **Step 1: 写测试 `CellConverterTest`**

```java
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import java.util.HashMap;
import java.util.Map;
import static org.assertj.core.api.Assertions.assertThat;

class CellConverterTest {
    private XSSFCell newCell() {
        XSSFWorkbook wb = new XSSFWorkbook();
        return wb.createSheet().createRow(0).createCell(0);
    }

    @Test
    void should_write_string_number_boolean() {
        CellConverter cc = new CellConverter(new StyleConverter(new XSSFWorkbook()));
        XSSFCell c1 = newCell(); cc.writeCell(c1, new ICellData().setV("A1").setT(CellValueType.STRING));
        XSSFCell c2 = newCell(); cc.writeCell(c2, new ICellData().setV(3.14).setT(CellValueType.NUMBER));
        XSSFCell c3 = newCell(); cc.writeCell(c3, new ICellData().setV(1).setT(CellValueType.BOOLEAN));
        assertThat(c1.getStringCellValue()).isEqualTo("A1");
        assertThat(c2.getNumericCellValue()).isEqualTo(3.14);
        assertThat(c3.getBooleanCellValue()).isTrue();
    }

    @Test
    void should_force_text_set_quote_prefix() {
        CellConverter cc = new CellConverter(new StyleConverter(new XSSFWorkbook()));
        XSSFCell c = newCell();
        cc.writeCell(c, new ICellData().setV("012.0").setT(CellValueType.FORCE_TEXT));
        assertThat(c.getCellStyle().getQuotePrefix()).isTrue();
        assertThat(c.getStringCellValue()).isEqualTo("012.0");
        ICellData back = cc.readCell(c);
        assertThat(back.getT()).isEqualTo(CellValueType.FORCE_TEXT);
    }

    @Test
    void should_write_formula_and_read_back() {
        CellConverter cc = new CellConverter(new StyleConverter(new XSSFWorkbook()));
        XSSFCell c = newCell();
        cc.writeCell(c, new ICellData().setF("SUM(A1:B1)").setV(3.0).setT(CellValueType.NUMBER));
        ICellData back = cc.readCell(c);
        assertThat(back.getF()).isEqualTo("SUM(A1:B1)");
    }

    @Test
    void should_preserve_custom_field_via_extras_on_cell() {
        // custom 无 xlsx 原生载体 → 调用方（WorkbookConverter）应放入边车；
        // CellConverter 在 readCell 结果里保留空 custom；写 Cell 时不落盘。
        CellConverter cc = new CellConverter(new StyleConverter(new XSSFWorkbook()));
        Map<String, Object> custom = new HashMap<>(); custom.put("k", "v");
        XSSFCell c = newCell();
        cc.writeCell(c, new ICellData().setV("x").setCustom(custom));
        assertThat(c.getStringCellValue()).isEqualTo("x");
    }
}
```

- [ ] **Step 2: 运行确认失败**

Run: `mvn -q test -Dtest=CellConverterTest`
Expected: 编译失败。

- [ ] **Step 3: 实现 `CellConverter`**

要点：
- `writeCell`：
  - `src.f != null` → `cell.setCellFormula(f)`，若同时提供 `v` 再按 `t` 设置 cachedValue。
  - `t=FORCE_TEXT`：`cellStyle.setQuotePrefix(true)` 并按字符串写。
  - `t=BOOLEAN`：`v` 可为 `Number(0/1)` 或 `Boolean`，统一归一化为 boolean。
  - `t=STRING` 或推断字符串：`cell.setCellValue(String.valueOf(v))`。
  - `t=NUMBER` 或数值：`cell.setCellValue(((Number)v).doubleValue())`。
  - `src.s` 若是 `IStyleData` 对象，调用 `styles.toPoiStyle` 应用；若是 `String`（styleId）→ 调用方先解引用，`CellConverter` 不处理。
- `readCell`：
  - `FORMULA` → `f = cell.getCellFormula()`，`v = cached`，`t` 按 cached 类型。
  - `STRING` 且 `cellStyle.getQuotePrefix()` → `t=FORCE_TEXT`。
  - `BOOLEAN` → `v = cell.getBooleanCellValue()?1:0`，`t=BOOLEAN`。
  - `NUMERIC` → `v = cell.getNumericCellValue()`，`t=NUMBER`。
  - `s = styles.styleIdOf(fromPoiStyle(cell.getCellStyle()))`（字符串 id，调用方维护全局 styles map）。

- [ ] **Step 4: Commit**

```bash
git add src
git commit -m "feat(converter): add CellConverter (value/type/formula/force-text)"
```

---

## Task 9: 共享公式注册器 SharedFormulaRegistry

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/SharedFormulaRegistry.java`
- Test: `src/test/java/io/github/autoffice/univer/converter/SharedFormulaRegistryTest.java`

**接口**：

```java
public final class SharedFormulaRegistry {
    public String registerRead(int row, int col, String formula);     // 读侧：同公式表达式分组，返回 si；首个作主格
    public Optional<String> masterFormulaOf(String si);                 // 主格公式
    public Map<String, int[]> masterCoordOf();                          // si → {row,col}
    public void registerWrite(int row, int col, String si, String f);   // 写侧：收集 (si, f) → 落盘时按 si 分组、主格置右下
    public void applyOnWorkbook(XSSFWorkbook wb);                       // 写侧：遍历分组，主格写 f，其余写 shared formula reference
}
```

- [ ] **Step 1: 写测试**

```java
package io.github.autoffice.univer.converter;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.util.Arrays;
import static org.assertj.core.api.Assertions.assertThat;

class SharedFormulaRegistryTest {

    @Test
    void should_group_same_formula_into_one_si() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        String s1 = r.registerRead(0, 0, "A$1+1");
        String s2 = r.registerRead(0, 1, "A$1+1");
        assertThat(s1).isEqualTo(s2);
        assertThat(r.masterFormulaOf(s1)).contains("A$1+1");
    }

    @Test
    void should_pick_bottom_right_as_master_on_write() {
        SharedFormulaRegistry r = new SharedFormulaRegistry();
        r.registerWrite(0, 0, "si-x", "SUM(A1:B1)");
        r.registerWrite(2, 3, "si-x", "SUM(A1:B1)");     // 右下
        r.registerWrite(1, 2, "si-x", "SUM(A1:B1)");
        XSSFWorkbook wb = new XSSFWorkbook(); wb.createSheet("s");
        r.applyOnWorkbook(wb);
        assertThat(wb.getSheetAt(0).getRow(2).getCell(3).getCellFormula()).isEqualTo("SUM(A1:B1)");
    }
}
```

- [ ] **Step 2: 实现**（关键算法：写侧按 `si` 分组，求 `max(row,col)` 作为主格；主格 `setCellFormula(f)`；其余 cell 使用 POI `CTCellFormula` 配置 `t=shared` 与 `si` 索引；读侧遇到 POI shared formula 拆表达式回原式，再按「表达式+主格坐标」生成 `si`）。

- [ ] **Step 3: 运行测试并 Commit**

```bash
mvn -q test -Dtest=SharedFormulaRegistryTest
git add src
git commit -m "feat(converter): add SharedFormulaRegistry for si grouping"
```

---

## Task 10: 富文本转换器 RichTextConverter

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/RichTextConverter.java`
- Test: `src/test/java/io/github/autoffice/univer/converter/RichTextConverterTest.java`

**接口**：

```java
public final class RichTextConverter {
    public RichTextConverter(XSSFWorkbook wb);
    public XSSFRichTextString toPoi(IDocumentData p);
    public IDocumentData fromPoi(XSSFRichTextString rts);
}
```

- [ ] **Step 1: 测试：一段富文本含 2 个 run（Arial 10 黑色 / Arial 14 红色）+ 段落换行**

```java
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import java.util.Arrays;
import static org.assertj.core.api.Assertions.assertThat;

class RichTextConverterTest {

    @Test
    void should_roundtrip_runs_and_paragraphs() {
        XSSFWorkbook wb = new XSSFWorkbook();
        RichTextConverter rc = new RichTextConverter(wb);
        IDocumentData.TextRun run1 = new IDocumentData.TextRun().setSt(0).setEd(5)
                .setTs(new IStyleData().setFf("Arial").setFs(10));
        IDocumentData.TextRun run2 = new IDocumentData.TextRun().setSt(6).setEd(11)
                .setTs(new IStyleData().setFf("Arial").setFs(14).setCl(new IColorStyle().setRgb("#ff0000")));
        IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("hello\nworld")
                .setTextRuns(Arrays.asList(run1, run2))
                .setParagraphs(Arrays.asList(new IDocumentData.Paragraph().setStartIndex(5))));
        XSSFRichTextString rts = rc.toPoi(p);
        assertThat(rts.getString()).contains("hello").contains("world");
        IDocumentData back = rc.fromPoi(rts);
        assertThat(back.getBody().getDataStream()).isEqualTo("hello\nworld");
        assertThat(back.getBody().getTextRuns()).hasSize(2);
    }
}
```

- [ ] **Step 2: 实现**

要点：
- 写侧：`new XSSFRichTextString(dataStream)` 后按 textRun 的 `[st, ed)` 区间调 `applyFont(st, ed, font)`。
- 读侧：通过 `rts.numFormattingRuns()` + `rts.getIndexOfFormattingRun(i)` / `rts.getLengthOfFormattingRun(i)` 切 run；font 通过 `rts.getFontOfFormattingRun(i)` 还原 `IStyleData`（只填字体相关字段）。
- `paragraphs.startIndex` 简单由 `\n` 位置回填。

- [ ] **Step 3: 测试 + Commit**

```bash
mvn -q test -Dtest=RichTextConverterTest
git add src
git commit -m "feat(converter): add RichTextConverter for IDocumentData <-> XSSFRichTextString"
```

---

## Task 11: 边车分区 SidecarPart

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/resource/SidecarPart.java`
- Test: `src/test/java/io/github/autoffice/univer/resource/SidecarPartTest.java`

**接口**：

```java
public final class SidecarPart {
    public static final String PART_NAME = "/univer/metadata.json";
    public static final String CONTENT_TYPE = "application/json";
    public static final String REL_TYPE = "http://schemas.autoffice.io/univer/2026/metadata";

    public static Optional<IWorkbookData> read(OPCPackage pkg);
    public static void write(OPCPackage pkg, IWorkbookData wb, boolean pretty);
}
```

- [ ] **Step 1: 测试：写入 → 关闭 → 重新打开读回，内容一致；无边车时返回 empty**

```java
package io.github.autoffice.univer.resource;

import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.util.Optional;
import static org.assertj.core.api.Assertions.assertThat;

class SidecarPartTest {
    @Test
    void should_roundtrip_sidecar_part() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s1");
            IWorkbookData meta = new IWorkbookData().setId("w").setAppVersion("0.10.2");
            meta.getSheets().put("s1", new IWorksheetData().setId("s1").setName("s1"));
            SidecarPart.write(wb.getPackage(), meta, false);
            wb.write(buf);
        }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            Optional<IWorkbookData> read = SidecarPart.read(pkg);
            assertThat(read).isPresent();
            assertThat(read.get().getAppVersion()).isEqualTo("0.10.2");
        }
    }

    @Test
    void should_return_empty_when_sidecar_missing() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) { wb.createSheet("s1"); wb.write(buf); }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            assertThat(SidecarPart.read(pkg)).isEmpty();
        }
    }
}
```

- [ ] **Step 2: 实现**

要点：
- 使用 `PackagePartName partName = PackagingURIHelper.createPartName(PART_NAME);`。
- 创建/查询：`pkg.getPart(partName)`；不存在则 `pkg.createPart(partName, CONTENT_TYPE)`。
- 关系：在根关系（`pkg.addRelationship`）添加 `targetURI=partName.getURI()`, `targetMode=INTERNAL`, `relationshipType=REL_TYPE`。
- JSON：`JsonMapper.get().writerWithDefaultPrettyPrinter()`（pretty）或 `writer()`，写入 `part.getOutputStream()`。

- [ ] **Step 3: 测试 + Commit**

```bash
mvn -q test -Dtest=SidecarPartTest
git add src
git commit -m "feat(resource): add OPC sidecar part read/write"
```

---

## Task 12: 工作表转换器 WorksheetConverter

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/WorksheetConverter.java`
- Test: `src/test/java/io/github/autoffice/univer/converter/WorksheetConverterTest.java`

**接口**：

```java
public final class WorksheetConverter {
    public WorksheetConverter(XSSFWorkbook wb, StyleConverter styles, CellConverter cells,
                              RichTextConverter rich, SharedFormulaRegistry formulas,
                              UniverXlsxOptions opts);
    public void writeSheet(XSSFSheet sheet, IWorksheetData src);
    public IWorksheetData readSheet(XSSFSheet sheet);
}
```

- [ ] **Step 1: 测试覆盖 mergeData / freeze / hidden rows / column width / showGridlines / rightToLeft / zoomRatio**

```java
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import java.util.Arrays;
import java.util.LinkedHashMap;
import static org.assertj.core.api.Assertions.assertThat;

class WorksheetConverterTest {
    @Test
    void should_roundtrip_sheet_attributes_and_cells() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            CellConverter cc = new CellConverter(sc);
            RichTextConverter rc = new RichTextConverter(wb);
            SharedFormulaRegistry sfr = new SharedFormulaRegistry();
            WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());

            IWorksheetData src = new IWorksheetData().setId("s1").setName("Sheet1")
                .setFreeze(new IFreeze().setStartRow(1).setStartColumn(1).setXSplit(1).setYSplit(1))
                .setShowGridlines(BooleanNumber.FALSE)
                .setRightToLeft(BooleanNumber.TRUE)
                .setZoomRatio(1.25)
                .setMergeData(Arrays.asList(new IRange().setStartRow(0).setStartColumn(0).setEndRow(1).setEndColumn(1)));
            LinkedHashMap<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("A1").setT(CellValueType.STRING));
            src.getCellData().put(0, row);

            XSSFSheet sheet = wb.createSheet("Sheet1");
            wsc.writeSheet(sheet, src);
            assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("A1:B2");
            assertThat(sheet.isDisplayGridlines()).isFalse();
            assertThat(sheet.isRightToLeft()).isTrue();

            IWorksheetData back = wsc.readSheet(sheet);
            assertThat(back.getMergeData()).hasSize(1);
            assertThat(back.getShowGridlines()).isEqualTo(BooleanNumber.FALSE);
            assertThat(back.getCellData().get(0).get(0).getV()).isEqualTo("A1");
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
```

- [ ] **Step 2: 实现**

要点：
- `writeSheet`：按 `src.cellData` 的稀疏结构遍历 → `sheet.createRow(r).createCell(c)` → `CellConverter.writeCell`。
- `mergeData` → `sheet.addMergedRegion(new CellRangeAddress(startRow,endRow,startColumn,endColumn))`。
- `freeze`：`sheet.createFreezePane(xSplit, ySplit, startColumn, startRow)`；坐标为 `-1` 时跳过。
- `defaultColumnWidth = LengthUtils.pxToChars(src.defaultColumnWidth)`；`defaultRowHeightInPoints = LengthUtils.pxToPoints(src.defaultRowHeight)`。
- `rowData[i].h` → `row.setHeightInPoints(pxToPoints(h))`；`hd=1` → `setZeroHeight(true)`。
- `columnData[j].w` → `sheet.setColumnWidth(j, charsToPx→POI 内部单位: (int)(chars*256))`；`hd` → `setColumnHidden`。
- `showGridlines/rightToLeft/zoomRatio` 直接写 POI 对应属性。
- `readSheet`：对称读取；`rowCount/columnCount` 从 sheet 实际使用范围取（边车会覆盖）。

- [ ] **Step 3: 测试 + Commit**

```bash
mvn -q test -Dtest=WorksheetConverterTest
git add src
git commit -m "feat(converter): add WorksheetConverter for sheet attributes and cells"
```

---

## Task 13: 工作簿转换器 WorkbookConverter + Reader/Writer 装配

**Files:**
- Create: `src/main/java/io/github/autoffice/univer/converter/WorkbookConverter.java`
- Modify: `src/main/java/io/github/autoffice/univer/io/UniverXlsxReader.java`
- Modify: `src/main/java/io/github/autoffice/univer/io/UniverXlsxWriter.java`
- Test: `src/test/java/io/github/autoffice/univer/io/RoundTripSmokeTest.java`

- [ ] **Step 1: 实现 `WorkbookConverter`**

```java
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.*;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

/** 工作簿级双向转换 / workbook-level converter. */
public final class WorkbookConverter {
    private final UniverXlsxOptions opts;
    public WorkbookConverter(UniverXlsxOptions opts) { this.opts = opts; }

    public XSSFWorkbook toXlsx(IWorkbookData src) {
        XSSFWorkbook wb = new XSSFWorkbook();
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, opts);

        List<String> order = src.getSheetOrder() != null ? src.getSheetOrder()
            : new ArrayList<>(src.getSheets().keySet());
        for (String sid : order) {
            IWorksheetData ws = resolveStyles(src, src.getSheets().get(sid));
            XSSFSheet sheet = wb.createSheet(ws.getName() == null ? sid : ws.getName());
            wsc.writeSheet(sheet, ws);
            if (ws.getHidden() == BooleanNumber.TRUE) {
                wb.setSheetVisibility(wb.getSheetIndex(sheet), SheetVisibility.HIDDEN);
            }
        }
        sfr.applyOnWorkbook(wb);
        return wb;
    }

    public IWorkbookData fromXlsx(XSSFWorkbook wb, IWorkbookData sidecarBaseline) {
        IWorkbookData out = sidecarBaseline != null ? sidecarBaseline : new IWorkbookData();
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, opts);

        List<String> order = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            IWorksheetData ws = wsc.readSheet(sheet);
            String sid = ws.getId() != null ? ws.getId() : sheet.getSheetName();
            ws.setId(sid);
            if (wb.getSheetVisibility(i) == SheetVisibility.HIDDEN) ws.setHidden(BooleanNumber.TRUE);
            out.getSheets().putIfAbsent(sid, ws);
            order.add(sid);
        }
        if (out.getSheetOrder() == null) out.setSheetOrder(order);
        if (out.getLocale() == null) out.setLocale(opts.getLocale());
        return out;
    }

    private IWorksheetData resolveStyles(IWorkbookData wb, IWorksheetData ws) {
        if (wb.getStyles() == null || ws == null || ws.getCellData() == null) return ws;
        ws.getCellData().forEach((r, cols) -> cols.forEach((c, cd) -> {
            if (cd != null && cd.getS() instanceof String) {
                IStyleData inline = wb.getStyles().get((String) cd.getS());
                if (inline != null) cd.setS(inline);
            }
        }));
        return ws;
    }
}
```

- [ ] **Step 2: 实现 `UniverXlsxReader`**

```java
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.*;
import io.github.autoffice.univer.converter.WorkbookConverter;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.resource.SidecarPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;

/** xlsx → IWorkbookData / xlsx reader. */
public final class UniverXlsxReader {
    private final UniverXlsxOptions opts;
    public UniverXlsxReader(UniverXlsxOptions opts) { this.opts = opts; }

    public IWorkbookData read(InputStream in) throws UniverXlsxException {
        try (OPCPackage pkg = OPCPackage.open(in)) {
            Optional<IWorkbookData> sidecar = SidecarPart.read(pkg);
            try (XSSFWorkbook wb = new XSSFWorkbook(pkg)) {
                return new WorkbookConverter(opts).fromXlsx(wb, sidecar.orElse(null));
            }
        } catch (IOException e) {
            throw new UniverXlsxReadException("read xlsx failed", e);
        } catch (Exception e) {
            throw new UniverXlsxReadException("read xlsx failed: " + e.getMessage(), e);
        }
    }
}
```

- [ ] **Step 3: 实现 `UniverXlsxWriter`**

```java
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.*;
import io.github.autoffice.univer.converter.WorkbookConverter;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.resource.SidecarPart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

/** IWorkbookData → xlsx / xlsx writer. */
public final class UniverXlsxWriter {
    private final UniverXlsxOptions opts;
    public UniverXlsxWriter(UniverXlsxOptions opts) { this.opts = opts; }

    public void write(IWorkbookData src, OutputStream out) throws UniverXlsxException {
        try (XSSFWorkbook wb = new WorkbookConverter(opts).toXlsx(src)) {
            if (opts.isWriteSidecar()) SidecarPart.write(wb.getPackage(), src, opts.isPrettyJson());
            wb.write(out);
        } catch (IOException e) {
            throw new UniverXlsxWriteException("write xlsx failed", e);
        }
    }
}
```

- [ ] **Step 4: 写 `RoundTripSmokeTest`**

```java
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.*;
import org.junit.jupiter.api.Test;
import java.io.*;
import java.util.*;
import static org.assertj.core.api.Assertions.assertThat;

class RoundTripSmokeTest {
    @Test
    void should_roundtrip_minimal_workbook() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("w").setAppVersion("0.10.2").setLocale("zhCN");
        src.setSheetOrder(Collections.singletonList("s1"));
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("Sheet1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("A1").setT(CellValueType.STRING));
        row.put(1, new ICellData().setV(3.14).setT(CellValueType.NUMBER));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));
        assertThat(back.getSheetOrder()).containsExactly("s1");
        assertThat(back.getSheets().get("s1").getCellData().get(0).get(0).getV()).isEqualTo("A1");
        assertThat(back.getAppVersion()).isEqualTo("0.10.2");
    }
}
```

- [ ] **Step 5: 测试 + Commit**

```bash
mvn -q test
git add src
git commit -m "feat(io): wire WorkbookConverter into reader/writer with sidecar round-trip"
```

---

## Task 14: 扩展 round-trip 测试（多 sheet / 样式 / 公式 / 富文本 / resources / custom / 严格模式）

**Files:**
- Create: `src/test/java/io/github/autoffice/univer/io/FullRoundTripTest.java`

- [ ] **Step 1: 写一组完整 round-trip 测试**

覆盖：
  1. 多 sheet（3 个，其中 1 个 hidden，sheetOrder 乱序）
  2. 全字段样式 + styles 去重（构造 3 个单元格引用同一 styleId，断言读回后 styles 中仅一份）
  3. 共享公式 si（4 个 cell 同 si，主格右下）
  4. 富文本 `p`（2 run + 段落）
  5. `mergeData`（交叉 + 行合并 + 列合并）
  6. 冻结三形态 + 无冻结
  7. `resources`（任意 JSON Map）round-trip 一致
  8. cell `custom` 字段 round-trip 一致
  9. `extras` 兜底：向 ICellData JSON 注入未知字段 `futureX`，断言 round-trip 后仍存在
 10. 外部 xlsx 兼容：用 POI 裸写一个 xlsx（无边车），本库读取不抛错，边车缺失字段走 fallback
 11. 严格模式：构造一个 `pd.t=5`（无 xlsx 原生载体），`strictMode=true` 时抛 `UniverXlsxUnsupportedFeatureException`
 12. 空 workbook（0 sheet 与 1 空 sheet 两种），round-trip 不抛错

每个用例写成独立 `@Test` 方法，命名 `should_xxx_when_yyy`。断言采用 `assertThat(actual).usingRecursiveComparison().ignoringFields("extras").isEqualTo(expected)`。

- [ ] **Step 2: 如果某些字段（如 `pd`）未在 StyleConverter 实现严格模式分支，补齐实现（落边车时记录 WARN 日志；严格模式下抛异常）**

- [ ] **Step 3: 运行全部测试**

Run: `mvn -q test`
Expected: 全绿。

- [ ] **Step 4: Commit**

```bash
git add src
git commit -m "test: full round-trip coverage (multi-sheet, style dedup, si, rich text, strict mode)"
```

---

## Task 15: 覆盖率与规范检查

- [ ] **Step 1: 运行 JaCoCo**

Run: `mvn -q -Pcoverage verify`
Expected: `BUILD SUCCESS`，查看 `target/site/jacoco/index.html` 行覆盖率 ≥ 80%；不足则补测试再跑。

- [ ] **Step 2: 运行 p3c 阿里规范检查**

Run: `mvn -q -Plint verify`
Expected: 无 PMD 错误（警告允许）；有错误按提示修正命名、Javadoc 等。

- [ ] **Step 3: 补 README**

Create: `README.md`（中英双语，含快速开始、安装 Maven 坐标、典型 read/write 示例、limitations 清单）。

- [ ] **Step 4: Commit**

```bash
git add pom.xml README.md
git commit -m "chore: enable coverage/lint profiles and add README"
```

---

## Task 16: 发布准备（可选）

- [ ] **Step 1: `pom.xml` 补 distributionManagement / licenses / scm / developers 信息**（面向 Sonatype OSSRH）
- [ ] **Step 2: 生成 sources jar 与 javadoc jar**：接入 `maven-source-plugin` 与 `maven-javadoc-plugin`，运行 `mvn -q package` 验证 `target/*-sources.jar` / `*-javadoc.jar`
- [ ] **Step 3: 打 tag**：`git tag v0.1.0` 并 commit `chore: prepare 0.1.0 release artifacts`

---

## Self-Review

- **Spec coverage**：Task 2–4 覆盖所有 POJO；Task 5 覆盖工具类；Task 7 覆盖样式全字段映射；Task 8 覆盖单元格值/类型/强文本/公式；Task 9 覆盖 si；Task 10 覆盖富文本；Task 11 覆盖边车；Task 12 覆盖工作表级字段；Task 13 装配 Reader/Writer 并覆盖 styles 解引用、hidden sheet、sheetOrder；Task 14 覆盖多 sheet/custom/resources/extras/外部兼容/严格模式/空簿；Task 15 覆盖覆盖率与 p3c。Spec 第 4、5、6 节所有字段族均有任务承接。
- **Placeholder scan**：无 TBD / TODO / “implement later” 等字样；所有代码步骤都给了完整代码或 minutely 要点 + 接口签名。
- **Type consistency**：`StyleConverter / CellConverter / RichTextConverter / SharedFormulaRegistry / WorksheetConverter / WorkbookConverter` 签名在首次引入时给定，后续 Task 严格复用。
- **风险**：`IntegerKeyDeserializer` 全局注册是否会误伤其他 `Map<Integer,?>`？因为库只处理自己的 POJO 且默认 key 反序列化器只作用于 JSON 对象键场景，风险可控；若发生冲突，Task 4 Step 8 允许回退到字段级 `@JsonDeserialize(keyUsing=...)`。






