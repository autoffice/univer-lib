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
 * 条件格式双向转换：构造 Univer 规则 → 写入 POI → 回读 → 核对核心字段。
 * Bidirectional CF conversion: construct Univer rule → POI → round-trip → assert core fields.
 */
class ConditionalFormattingConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    // ---------------------------------------------------------
    // highlightCell / formula
    // ---------------------------------------------------------
    @Test
    void should_roundtrip_highlightCell_formula() throws Exception {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "hl1");
        rule.set("ranges", oneRange(0, 0, 9, 4));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        body.put("value", "=MOD(COLUMN(B1),2)=0");
        ObjectNode style = mapper.createObjectNode();
        ObjectNode bg = mapper.createObjectNode();
        bg.put("rgb", "#cccccc");
        style.set("bg", bg);
        ObjectNode cl = mapper.createObjectNode();
        cl.put("rgb", "#35322b");
        style.set("cl", cl);
        body.set("style", style);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        ArrayNode arr = mapper.createArrayNode();
        arr.add(rule);

        ArrayNode roundTrip = writeReadBack(arr);

        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode rt = roundTrip.get(0);
        assertThat(rt.path("rule").path("type").asText()).isEqualTo("highlightCell");
        assertThat(rt.path("rule").path("subType").asText()).isEqualTo("formula");
        assertThat(rt.path("rule").path("value").asText()).contains("MOD");
        JsonNode rng = rt.path("ranges").get(0);
        assertThat(rng.path("startRow").asInt()).isEqualTo(0);
        assertThat(rng.path("endRow").asInt()).isEqualTo(9);
        assertThat(rng.path("startColumn").asInt()).isEqualTo(0);
        assertThat(rng.path("endColumn").asInt()).isEqualTo(4);
        assertThat(rng.path("rangeType").asInt()).isEqualTo(0);
        // 背景色 round-trip 应保留
        assertThat(rt.path("rule").path("style").path("bg").path("rgb").asText().toLowerCase())
                .isEqualTo("#cccccc");
    }

    // ---------------------------------------------------------
    // colorScale
    // ---------------------------------------------------------
    @Test
    void should_roundtrip_colorScale() throws Exception {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "cs1");
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "colorScale");
        ArrayNode config = mapper.createArrayNode();
        ObjectNode p0 = mapper.createObjectNode();
        p0.put("color", "#F1EAFA");
        p0.set("value", typedValue("min", 0));
        p0.put("index", 0);
        ObjectNode p1 = mapper.createObjectNode();
        p1.put("color", "#6721CB");
        p1.set("value", typedValue("max", 100));
        p1.put("index", 1);
        config.add(p0);
        config.add(p1);
        body.set("config", config);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        ArrayNode arr = mapper.createArrayNode();
        arr.add(rule);

        ArrayNode roundTrip = writeReadBack(arr);
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode rt = roundTrip.get(0);
        assertThat(rt.path("rule").path("type").asText()).isEqualTo("colorScale");
        JsonNode cfg = rt.path("rule").path("config");
        assertThat(cfg.isArray()).isTrue();
        assertThat(cfg.size()).isEqualTo(2);
        assertThat(cfg.get(0).path("value").path("type").asText()).isEqualTo("min");
        assertThat(cfg.get(1).path("value").path("type").asText()).isEqualTo("max");
        assertThat(cfg.get(0).path("color").asText().toLowerCase()).isEqualTo("#f1eafa");
        assertThat(cfg.get(1).path("color").asText().toLowerCase()).isEqualTo("#6721cb");
    }

    // ---------------------------------------------------------
    // dataBar
    // ---------------------------------------------------------
    @Test
    void should_roundtrip_dataBar() throws Exception {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "db1");
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        config.set("min", typedValue("min", 0));
        config.set("max", typedValue("max", 100));
        config.put("isGradient", true);
        config.put("positiveColor", "#abd91a");
        config.put("nativeColor", "#ffbe38");
        body.set("config", config);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        ArrayNode arr = mapper.createArrayNode();
        arr.add(rule);

        ArrayNode roundTrip = writeReadBack(arr);
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode rt = roundTrip.get(0);
        assertThat(rt.path("rule").path("type").asText()).isEqualTo("dataBar");
        JsonNode cfg = rt.path("rule").path("config");
        assertThat(cfg.path("min").path("type").asText()).isEqualTo("min");
        assertThat(cfg.path("max").path("type").asText()).isEqualTo("max");
        assertThat(cfg.path("positiveColor").asText().toLowerCase()).isEqualTo("#abd91a");
    }

    // ---------------------------------------------------------
    // iconSet
    // ---------------------------------------------------------
    @Test
    void should_roundtrip_iconSet() throws Exception {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "is1");
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "iconSet");
        body.put("isShowValue", true);
        ArrayNode config = mapper.createArrayNode();
        String iconType = "3Arrows";
        config.add(iconStop("greaterThan", 300, iconType, "0"));
        config.add(iconStop("greaterThan", 100, iconType, "1"));
        config.add(iconStop("lessThanOrEqual", 50, iconType, "2"));
        body.set("config", config);
        rule.set("rule", body);
        rule.put("stopIfTrue", false);
        ArrayNode arr = mapper.createArrayNode();
        arr.add(rule);

        ArrayNode roundTrip = writeReadBack(arr);
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode rt = roundTrip.get(0);
        assertThat(rt.path("rule").path("type").asText()).isEqualTo("iconSet");
        JsonNode cfg = rt.path("rule").path("config");
        assertThat(cfg.isArray()).isTrue();
        assertThat(cfg.size()).isEqualTo(3);
        // Univer 侧 iconType 应保留
        assertThat(cfg.get(0).path("iconType").asText()).isEqualTo(iconType);
    }

    // ---------------------------------------------------------
    // helpers
    // ---------------------------------------------------------

    /**
     * Write the given Univer rules to a blank XSSFWorkbook, persist to bytes,
     * reopen with POI and read CF back to Univer JSON.
     */
    private ArrayNode writeReadBack(ArrayNode rules) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
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

    private ArrayNode oneRange(int startRow, int startCol, int endRow, int endCol) {
        ArrayNode arr = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", startRow);
        r.put("startColumn", startCol);
        r.put("endRow", endRow);
        r.put("endColumn", endCol);
        r.put("startAbsoluteRefType", 0);
        r.put("endAbsoluteRefType", 0);
        r.put("rangeType", 0);
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
