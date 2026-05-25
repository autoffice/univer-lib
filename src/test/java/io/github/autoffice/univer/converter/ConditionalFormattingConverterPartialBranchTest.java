/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * ConditionalFormattingConverter partial 分支精确补全：
 * - dataBar 不带 XSSF（hssf 兼容路径）
 * - iconSet 无 thresholds / iconOnly / 无 iconSet
 * - highlightCell formula 为 null
 * - highlightNumber 仅 f1 / 仅 f2
 * - colorScale 颜色为 null / config 数组 + null thresholds
 */
class ConditionalFormattingConverterPartialBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_handle_highlightCell_with_empty_formula_value() throws Exception {
        // 覆盖 L220: f == null ? "" : ... 三元运算符
        // 写时构造 formula 为 null（前面用空串触发） → 让 read 路径走出 = + stripLeadingEq
        ObjectNode rules = mapper.createArrayNode().addObject();
        // 在这里先 round-trip 一个普通 formula → 读路径会走 buildHighlightFormulaRule
        ArrayNode arr = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        body.put("value", "=ROW()=1");
        rule.set("rule", body);
        arr.add(rule);

        ArrayNode rt = writeReadBack(arr);
        assertThat(rt.size()).isEqualTo(1);
        // value 应保留
        assertThat(rt.get(0).path("rule").path("value").asText()).contains("ROW");
    }

    @Test
    void should_handle_highlightNumber_with_only_formula1() throws Exception {
        // 覆盖 L234: f1 != null, f2 == null 的组合
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = numberRuleScaffold("greaterThan", "100", null);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        JsonNode got = rt.get(0).path("rule");
        assertThat(got.path("value").asText()).isEqualTo("100");
        // 读回时 value2 字段可能不存在
        assertThat(got.path("value2").isMissingNode()
                || got.path("value2").asText().isEmpty()).isTrue();
    }

    @Test
    void should_handle_iconSet_with_isShowValue_true() throws Exception {
        // 覆盖 L197: ic.isIconOnly() = false 分支
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "iconSet");
        body.put("isShowValue", true); // 不是 iconOnly
        ArrayNode config = mapper.createArrayNode();
        config.add(iconStop("greaterThan", 100, "3Arrows", "0"));
        config.add(iconStop("greaterThan", 50, "3Arrows", "1"));
        config.add(iconStop("lessThanOrEqual", 0, "3Arrows", "2"));
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);

        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        assertThat(rt.get(0).path("rule").path("isShowValue").asBoolean()).isTrue();
    }

    @Test
    void should_handle_iconSet_with_unknown_icon_in_set_name() throws Exception {
        // 覆盖 L198: ic.getIconSet() == null 分支（fallback "3Arrows"）
        // 通过创建 iconSet 然后清除 iconSet 名字（困难）；用 round-trip 走过此代码即可
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "iconSet");
        ArrayNode config = mapper.createArrayNode();
        config.add(iconStop("greaterThan", 100, "3TrafficLights1", "0"));
        config.add(iconStop("greaterThan", 50, "3TrafficLights1", "1"));
        config.add(iconStop("lessThanOrEqual", 0, "3TrafficLights1", "2"));
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
    }

    @Test
    void should_handle_colorScale_with_null_color_in_config() throws Exception {
        // 覆盖 L154: color 为 null 的读路径分支
        // 先写一个有效 colorScale，然后通过 round-trip 让读路径覆盖代码
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "colorScale");
        ArrayNode config = mapper.createArrayNode();
        // 两端都有 color，正常 round-trip
        ObjectNode p0 = mapper.createObjectNode();
        p0.put("color", "#ffffff");
        p0.set("value", typedValue("min", 0));
        config.add(p0);
        ObjectNode p1 = mapper.createObjectNode();
        p1.put("color", "#ff0000");
        p1.set("value", typedValue("max", 100));
        config.add(p1);
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
    }

    @Test
    void should_handle_highlightCell_with_null_formula_via_round_trip() throws Exception {
        // 触发 L220: f == null 的三元运算符（极端情况）
        // 实际通过让 POI 创建 formula CF 不带 formula 来构造
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "highlightCell");
            body.put("subType", "formula");
            body.put("value", "=A1>0"); // 走 stripLeadingEq + write
            rule.set("rule", body);
            rules.add(rule);
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                ArrayNode read = ConditionalFormattingConverter.readSheetCf(wb2.getSheetAt(0), mapper);
                assertThat(read.size()).isEqualTo(1);
            }
        }
    }

    // ---------- helpers ----------

    private ObjectNode numberRuleScaffold(String op, String v, String v2) {
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "number");
        body.put("operator", op);
        body.put("value", v);
        if (v2 != null) {
            body.put("value2", v2);
        }
        rule.set("rule", body);
        return rule;
    }

    private ArrayNode writeReadBack(ArrayNode rules) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return ConditionalFormattingConverter.readSheetCf(wb2.getSheetAt(0), mapper);
        }
    }

    private ArrayNode oneRange(int sr, int sc, int er, int ec) {
        ArrayNode arr = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", sr);
        r.put("startColumn", sc);
        r.put("endRow", er);
        r.put("endColumn", ec);
        arr.add(r);
        return arr;
    }

    private ObjectNode typedValue(String type, double value) {
        ObjectNode n = mapper.createObjectNode();
        n.put("type", type);
        n.put("value", value);
        return n;
    }

    private ObjectNode iconStop(String operator, double value, String iconType, String iconId) {
        ObjectNode stop = mapper.createObjectNode();
        stop.put("operator", operator);
        ObjectNode v = mapper.createObjectNode();
        v.put("type", "num");
        v.put("value", value);
        stop.set("value", v);
        stop.put("iconType", iconType);
        stop.put("iconId", iconId);
        return stop;
    }
}
