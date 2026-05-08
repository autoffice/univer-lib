package io.github.autoffice.univer.io;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 用真实 xlsx 资源做端到端转换与回写校验。
 * End-to-end conversion test using a real xlsx fixture.
 *
 * <p>流程：test.xlsx → IWorkbookData → JSON 文件 → 读回 IWorkbookData → output.xlsx →
 * 再读 output.xlsx → 与原 IWorkbookData 比较；用于回归保证 round-trip 不丢数据。
 */
class TestXlsxRoundTripTest {

    private static final String FIXTURE = "/test.xlsx";

    @Test
    void should_roundtrip_real_xlsx_via_json_and_back() throws Exception {
        // 产物目录：target/roundtrip-output/，方便人工检查；mvn clean 时自动清理。
        Path tmp = Paths.get("target/roundtrip-output");
        Files.createDirectories(tmp);

        // 1. xlsx → IWorkbookData
        IWorkbookData firstRead;
        try (InputStream in = getClass().getResourceAsStream(FIXTURE)) {
            assertThat(in).as("fixture %s 必须存在", FIXTURE).isNotNull();
            firstRead = UniverXlsx.read(in);
        }
        assertThat(firstRead).isNotNull();
        assertThat(firstRead.getSheets()).as("至少应有一个工作表").isNotEmpty();
        assertThat(firstRead.getSheetOrder()).isNotEmpty();

        // 1b. 注入 SHEET_CONDITIONAL_FORMATTING_PLUGIN 资源，模拟 Univer 端携带的条件格式，
        //     让 round-trip 验证 Conditional Format sheet 的规则能通过 resources 正确流转。
        ObjectMapper mapper = JsonMapper.get();
        IWorksheetData cfSheet = findSheetByName(firstRead, "Conditional Format");
        assertThat(cfSheet).as("fixture 必须包含 Conditional Format 工作表").isNotNull();
        String cfSheetId = cfSheet.getId();
        injectConditionalFormatting(firstRead, cfSheetId, mapper);

        // 2. IWorkbookData → JSON 文件
        ObjectMapper mapperOut = mapper;
        Path jsonFile = tmp.resolve("test.workbook.json");
        mapperOut.writerWithDefaultPrettyPrinter().writeValue(jsonFile.toFile(), firstRead);
        assertThat(Files.size(jsonFile)).as("导出 JSON 不应为空").isGreaterThan(0L);

        // 3. JSON 文件 → IWorkbookData（验证 JSON 自身可被反序列化）
        IWorkbookData fromJson = mapper.readValue(jsonFile.toFile(), IWorkbookData.class);
        assertThat(fromJson)
                .usingRecursiveComparison()
                .ignoringFields("extras")
                .isEqualTo(firstRead);

        // 4. IWorkbookData → output.xlsx
        Path outXlsx = tmp.resolve("output.xlsx");
        UniverXlsx.write(fromJson, outXlsx);
        assertThat(Files.size(outXlsx)).as("output.xlsx 不应为空").isGreaterThan(0L);

        // 5. output.xlsx → IWorkbookData，再与第一次读出的快照对比
        IWorkbookData secondRead = UniverXlsx.read(outXlsx);

        // 5a. 关键工作簿级字段一致
        assertThat(secondRead.getSheetOrder())
                .as("sheetOrder 必须保留")
                .containsExactlyElementsOf(firstRead.getSheetOrder());
        assertThat(secondRead.getSheets().keySet())
                .as("sheet id 集合必须一致（不应出现 shadow 条目）")
                .containsExactlyInAnyOrderElementsOf(firstRead.getSheets().keySet());

        // 5b. 每个工作表的核心数据一致
        for (Map.Entry<String, IWorksheetData> e : firstRead.getSheets().entrySet()) {
            String sid = e.getKey();
            IWorksheetData expected = e.getValue();
            IWorksheetData actual = secondRead.getSheets().get(sid);
            assertThat(actual).as("sheet %s 必须存在", sid).isNotNull();

            assertThat(actual.getName()).as("sheet %s name", sid).isEqualTo(expected.getName());

            assertCellDataEqual(sid, expected, actual);

            java.util.List<io.github.autoffice.univer.model.IRange> expectedMerges = expected.getMergeData() == null
                    ? java.util.Collections.emptyList()
                    : expected.getMergeData();
            java.util.List<io.github.autoffice.univer.model.IRange> actualMerges = actual.getMergeData() == null
                    ? java.util.Collections.emptyList()
                    : actual.getMergeData();
            assertThat(actualMerges)
                    .as("sheet %s mergeData", sid)
                    .usingRecursiveFieldByFieldElementComparatorIgnoringFields("extras")
                    .containsExactlyInAnyOrderElementsOf(expectedMerges);
        }

        // 5c. 关键字段（appVersion、locale）应保留
        assertThat(secondRead.getLocale()).isEqualTo(firstRead.getLocale());
        if (firstRead.getAppVersion() != null) {
            assertThat(secondRead.getAppVersion()).isEqualTo(firstRead.getAppVersion());
        }

        // 5d. 专项验证：Cell sheet 的 A6（row=5, col=0）是一个 inline rich text，必须被识别为 p
        // Specific check: "Cell" sheet A6 should be recognized as inline rich text (p field populated).
        // test.xlsx 中 A6 内容为 "Inline Style Cell"，三段不同样式：
        //   "Inline" 红色粗体 / "Style" 黑色粗体带中划线 / "Cell" 黑色粗体
        IWorksheetData cellSheetFirst = findSheetByName(firstRead, "Cell");
        assertThat(cellSheetFirst).as("Cell 工作表必须存在").isNotNull();
        ICellData a6First = cellSheetFirst.getCellData().get(5).get(0);
        assertThat(a6First).as("Cell A6 必须存在").isNotNull();
        assertThat(a6First.getP())
                .as("Cell A6 应为 rich text（p 字段非空）")
                .isNotNull();
        assertThat(a6First.getP().getBody().getDataStream())
                .as("Cell A6 rich text 内容应为 'Inline Style Cell' + Univer 段落终止符")
                .isEqualTo("Inline Style Cell\r\n");
        assertThat(a6First.getP().getBody().getTextRuns())
                .as("Cell A6 应包含多个 text runs（红色粗体/黑色粗体中划线/黑色粗体）")
                .hasSizeGreaterThanOrEqualTo(2);

        // round-trip 后 Cell A6 同样应是 rich text
        IWorksheetData cellSheetSecond = findSheetByName(secondRead, "Cell");
        ICellData a6Second = cellSheetSecond.getCellData().get(5).get(0);
        assertThat(a6Second.getP())
                .as("round-trip 后 Cell A6 仍应为 rich text")
                .isNotNull();
        assertThat(a6Second.getP().getBody().getDataStream())
                .isEqualTo("Inline Style Cell\r\n");

        // 5e. 条件格式 round-trip：resources 中必须存在 SHEET_CONDITIONAL_FORMATTING_PLUGIN，
        //     且 Conditional Format sheet 对应的规则列表非空。
        JsonNode secondCfBySheet = extractCfResource(secondRead, mapper);
        assertThat(secondCfBySheet)
                .as("second read 必须包含 SHEET_CONDITIONAL_FORMATTING_PLUGIN 资源")
                .isNotNull();
        JsonNode cfRules = secondCfBySheet.get(cfSheetId);
        assertThat(cfRules)
                .as("Conditional Format sheet 的规则列表必须存在")
                .isNotNull();
        assertThat(cfRules.isArray()).isTrue();
        assertThat(cfRules.size()).as("CF 规则数量非空").isGreaterThan(0);
        // 打印片段便于人工核验
        String resPreview = mapper.writerWithDefaultPrettyPrinter()
                .writeValueAsString(secondRead.getResources());
        if (resPreview.length() > 1200) {
            resPreview = resPreview.substring(0, 1200) + "...";
        }
        System.out.println("[resources preview]\n" + resPreview);
    }

    /** 按 name 查找工作表（不区分 id 与 name 的差异）。 */
    private static IWorksheetData findSheetByName(IWorkbookData wb, String name) {
        for (IWorksheetData ws : wb.getSheets().values()) {
            if (name.equals(ws.getName())) {
                return ws;
            }
        }
        return null;
    }

    /** 比较两个工作表的 cellData，按值与类型逐格断言（忽略 extras）。 */
    private static void assertCellDataEqual(String sid, IWorksheetData expected, IWorksheetData actual) {
        Map<Integer, Map<Integer, ICellData>> exp = expected.getCellData();
        Map<Integer, Map<Integer, ICellData>> act = actual.getCellData();
        if (exp == null || exp.isEmpty()) {
            assertThat(act == null || act.isEmpty())
                    .as("sheet %s 期望空 cellData", sid)
                    .isTrue();
            return;
        }
        assertThat(act).as("sheet %s cellData 行集合", sid).containsKeys(exp.keySet().toArray(new Integer[0]));
        for (Map.Entry<Integer, Map<Integer, ICellData>> rowE : exp.entrySet()) {
            int row = rowE.getKey();
            Map<Integer, ICellData> expRow = rowE.getValue();
            Map<Integer, ICellData> actRow = act.get(row);
            assertThat(actRow).as("sheet %s row %d", sid, row).isNotNull();
            for (Map.Entry<Integer, ICellData> cellE : expRow.entrySet()) {
                int col = cellE.getKey();
                ICellData expCell = cellE.getValue();
                ICellData actCell = actRow.get(col);
                assertThat(actCell).as("sheet %s (%d,%d) 必须存在", sid, row, col).isNotNull();
                // 公式 cell 的 cached value 和类型可能在 round-trip 后变化（不做公式计算），跳过比较
                // Formula cells may have different cached values/types after round-trip, skip comparison
                if (expCell.getF() == null) {
                    assertThat(normalizeValue(actCell.getV()))
                            .as("sheet %s (%d,%d) 值", sid, row, col)
                            .isEqualTo(normalizeValue(expCell.getV()));
                    assertThat(actCell.getT())
                            .as("sheet %s (%d,%d) 类型", sid, row, col)
                            .isEqualTo(expCell.getT());
                }
                assertThat(actCell.getF())
                        .as("sheet %s (%d,%d) 公式", sid, row, col)
                        .isEqualTo(expCell.getF());
            }
        }
    }

    /** Jackson 反序列化数值 fall-back 到 Integer/Double，统一转 double 便于比较。 */
    private static Object normalizeValue(Object v) {
        if (v instanceof Number) {
            return ((Number) v).doubleValue();
        }
        return v;
    }

    /**
     * 注入 SHEET_CONDITIONAL_FORMATTING_PLUGIN 资源条目，供 round-trip 验证。
     * 构造三条规则：highlightCell/formula、colorScale、dataBar，覆盖多种主型。
     */
    @SuppressWarnings("unchecked")
    private static void injectConditionalFormatting(IWorkbookData wb, String sheetId, ObjectMapper mapper) throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        rules.add(buildHighlightRule(mapper));
        rules.add(buildColorScaleRule(mapper));
        rules.add(buildDataBarRule(mapper));

        ObjectNode bySheet = mapper.createObjectNode();
        bySheet.set(sheetId, rules);

        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_CONDITIONAL_FORMATTING_PLUGIN");
        entry.put("data", mapper.writeValueAsString(bySheet));

        List<Object> resources;
        Object existing = wb.getResources();
        if (existing instanceof List) {
            resources = (List<Object>) existing;
        } else {
            resources = new ArrayList<>();
        }
        resources.add(entry);
        wb.setResources(resources);
    }

    private static ObjectNode buildHighlightRule(ObjectMapper mapper) {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "testHl");
        ArrayNode ranges = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", 0);
        r.put("startColumn", 0);
        r.put("endRow", 9);
        r.put("endColumn", 4);
        r.put("startAbsoluteRefType", 0);
        r.put("endAbsoluteRefType", 0);
        r.put("rangeType", 0);
        ranges.add(r);
        rule.set("ranges", ranges);
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        body.put("value", "=MOD(ROW(A1),2)=0");
        ObjectNode style = mapper.createObjectNode();
        ObjectNode bg = mapper.createObjectNode();
        bg.put("rgb", "#cccccc");
        style.set("bg", bg);
        body.set("style", style);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        return rule;
    }

    private static ObjectNode buildColorScaleRule(ObjectMapper mapper) {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "testCs");
        ArrayNode ranges = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", 12);
        r.put("startColumn", 0);
        r.put("endRow", 21);
        r.put("endColumn", 0);
        r.put("startAbsoluteRefType", 0);
        r.put("endAbsoluteRefType", 0);
        r.put("rangeType", 0);
        ranges.add(r);
        rule.set("ranges", ranges);
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "colorScale");
        ArrayNode config = mapper.createArrayNode();
        ObjectNode p0 = mapper.createObjectNode();
        p0.put("color", "#f1eafa");
        ObjectNode v0 = mapper.createObjectNode();
        v0.put("type", "min");
        v0.put("value", 0);
        p0.set("value", v0);
        p0.put("index", 0);
        ObjectNode p1 = mapper.createObjectNode();
        p1.put("color", "#6721cb");
        ObjectNode v1 = mapper.createObjectNode();
        v1.put("type", "max");
        v1.put("value", 100);
        p1.set("value", v1);
        p1.put("index", 1);
        config.add(p0);
        config.add(p1);
        body.set("config", config);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        return rule;
    }

    private static ObjectNode buildDataBarRule(ObjectMapper mapper) {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "testDb");
        ArrayNode ranges = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", 24);
        r.put("startColumn", 0);
        r.put("endRow", 33);
        r.put("endColumn", 0);
        r.put("startAbsoluteRefType", 0);
        r.put("endAbsoluteRefType", 0);
        r.put("rangeType", 0);
        ranges.add(r);
        rule.set("ranges", ranges);
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        ObjectNode min = mapper.createObjectNode();
        min.put("type", "min");
        min.put("value", 0);
        config.set("min", min);
        ObjectNode max = mapper.createObjectNode();
        max.put("type", "max");
        max.put("value", 100);
        config.set("max", max);
        config.put("isGradient", true);
        config.put("positiveColor", "#abd91a");
        config.put("nativeColor", "#abd91a");
        body.set("config", config);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        return rule;
    }

    /** 从 resources 中取出 SHEET_CONDITIONAL_FORMATTING_PLUGIN 的 data（解析为 JsonNode）。 */
    @SuppressWarnings("unchecked")
    private static JsonNode extractCfResource(IWorkbookData wb, ObjectMapper mapper) throws Exception {
        Object res = wb.getResources();
        if (!(res instanceof List)) {
            return null;
        }
        for (Object item : (List<Object>) res) {
            if (!(item instanceof Map)) {
                continue;
            }
            Map<String, Object> m = (Map<String, Object>) item;
            if (!"SHEET_CONDITIONAL_FORMATTING_PLUGIN".equals(String.valueOf(m.get("name")))) {
                continue;
            }
            Object data = m.get("data");
            if (data instanceof String) {
                return mapper.readTree((String) data);
            }
            if (data != null) {
                return mapper.valueToTree(data);
            }
        }
        return null;
    }
}
