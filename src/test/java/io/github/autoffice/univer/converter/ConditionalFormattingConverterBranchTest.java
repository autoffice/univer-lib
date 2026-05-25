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
 * 补充 CF 转换器的写/读分支：number 各运算符、text、duplicate/unique、字体格式、stopIfTrue、空规则等。
 */
class ConditionalFormattingConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_handle_null_or_empty_rule_arrays() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ConditionalFormattingConverter.writeSheetCf(sh, null);
            ConditionalFormattingConverter.writeSheetCf(sh, mapper.createArrayNode());
            ConditionalFormattingConverter.writeSheetCf(null, mapper.createArrayNode());
            // readSheetCf null sheet
            assertThat(ConditionalFormattingConverter.readSheetCf(null, mapper)).isEmpty();
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_invalid_rule_nodes() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            rules.add(mapper.nullNode());
            // 没 ranges
            ObjectNode noRange = mapper.createObjectNode();
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "highlightCell");
            body.put("subType", "formula");
            body.put("value", "=A1=1");
            noRange.set("rule", body);
            rules.add(noRange);
            // 有 ranges 但 rule 为空
            ObjectNode noBody = mapper.createObjectNode();
            noBody.set("ranges", oneRange(0, 0, 0, 0));
            rules.add(noBody);
            // 未知 type
            ObjectNode unknownType = mapper.createObjectNode();
            unknownType.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode b = mapper.createObjectNode();
            b.put("type", "noSuchThing");
            unknownType.set("rule", b);
            rules.add(unknownType);

            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_highlight_formula_with_empty_value() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "highlightCell");
            body.put("subType", "formula");
            body.put("value", ""); // 空值 -> 跳过
            rule.set("rule", body);
            rules.add(rule);
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_roundtrip_highlight_number_with_between_operator() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = numberRuleScaffold("between", "10", "20");
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        JsonNode body = rt.get(0).path("rule");
        assertThat(body.path("subType").asText()).isEqualTo("number");
        assertThat(body.path("operator").asText()).isEqualTo("between");
        assertThat(body.path("value").asText()).isEqualTo("10");
        assertThat(body.path("value2").asText()).isEqualTo("20");
    }

    @Test
    void should_roundtrip_highlight_number_with_simple_operator_lt() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        rules.add(numberRuleScaffold("lessThan", "5", null));
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        JsonNode body = rt.get(0).path("rule");
        assertThat(body.path("operator").asText()).isEqualTo("lessThan");
    }

    @Test
    void should_roundtrip_highlight_number_with_each_simple_operator() throws Exception {
        String[] ops = {"equal", "notEqual", "greaterThan", "lessThan",
                        "greaterThanOrEqual", "lessThanOrEqual", "notBetween"};
        for (String op : ops) {
            ArrayNode rules = mapper.createArrayNode();
            rules.add(numberRuleScaffold(op, "1", "9"));
            ArrayNode rt = writeReadBack(rules);
            assertThat(rt.size()).as("op=%s", op).isEqualTo(1);
            assertThat(rt.get(0).path("rule").path("operator").asText()).isEqualTo(op);
        }
    }

    @Test
    void should_apply_font_formatting_in_highlight_rule() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        body.put("value", "=A1=\"hi\"");
        ObjectNode style = mapper.createObjectNode();
        ObjectNode cl = mapper.createObjectNode();
        cl.put("rgb", "#ff00ff");
        style.set("cl", cl);
        style.put("bl", 1);
        style.put("it", 1);
        ObjectNode st = mapper.createObjectNode();
        st.put("s", 1);
        style.set("st", st);
        body.set("style", style);
        rule.set("rule", body);
        rules.add(rule);

        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        JsonNode rtBody = rt.get(0).path("rule");
        assertThat(rtBody.path("style").path("bl").asInt()).isEqualTo(1);
        assertThat(rtBody.path("style").path("it").asInt()).isEqualTo(1);
        assertThat(rtBody.path("style").path("cl").path("rgb").asText().toLowerCase())
                .isEqualTo("#ff00ff");
    }

    @Test
    void should_handle_text_subtype_via_fallback_formula() throws Exception {
        // text/duplicate/unique 走 fallback 公式分支；写入不抛错即可
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "containsText");
        body.put("value", "x\"test");
        rule.set("rule", body);
        rules.add(rule);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_apply_stop_if_true_when_requested() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 0, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "highlightCell");
        body.put("subType", "formula");
        body.put("value", "=ROW()=1");
        rule.set("rule", body);
        rule.put("stopIfTrue", true);
        rules.add(rule);

        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        // 反射成功后 stopIfTrue 应为 true
        assertThat(rt.get(0).path("stopIfTrue").asBoolean()).isTrue();
    }

    @Test
    void should_handle_iconSet_with_unknown_iconType_falls_back() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "iconSet");
            body.put("isShowValue", false);
            ArrayNode config = mapper.createArrayNode();
            config.add(iconStop("greaterThan", 100, "weirdName", "0"));
            config.add(iconStop("greaterThan", 50, "weirdName", "1"));
            config.add(iconStop("lessThanOrEqual", 0, "weirdName", "2"));
            body.set("config", config);
            rule.set("rule", body);
            rules.add(rule);
            // 仅验证 writeSheetCf 不抛错且产生了规则；不做完整 read-back
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_iconSet_with_4_stops_fallback_to_GYR_4_arrows() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "iconSet");
            ArrayNode config = mapper.createArrayNode();
            config.add(iconStop("greaterThan", 75, "weirdName4", "0"));
            config.add(iconStop("greaterThan", 50, "weirdName4", "1"));
            config.add(iconStop("greaterThan", 25, "weirdName4", "2"));
            config.add(iconStop("lessThanOrEqual", 0, "weirdName4", "3"));
            body.set("config", config);
            rule.set("rule", body);
            rules.add(rule);
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_iconSet_with_5_stops_fallback() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "iconSet");
            ArrayNode config = mapper.createArrayNode();
            for (int i = 0; i < 5; i++) {
                config.add(iconStop("greaterThan", 80 - i * 20, "weirdName5", String.valueOf(i)));
            }
            body.set("config", config);
            rule.set("rule", body);
            rules.add(rule);
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_iconSet_empty_config() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ArrayNode rules = mapper.createArrayNode();
            ObjectNode rule = mapper.createObjectNode();
            rule.set("ranges", oneRange(0, 0, 0, 0));
            ObjectNode body = mapper.createObjectNode();
            body.put("type", "iconSet");
            body.set("config", mapper.createArrayNode());
            rule.set("rule", body);
            rules.add(rule);
            ConditionalFormattingConverter.writeSheetCf(sh, rules);
            assertThat(sh.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_threshold_with_percent_formula_and_string_value() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.put("cfId", "cs2");
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "colorScale");
        ArrayNode config = mapper.createArrayNode();
        // percent / percentile / formula / num 各种类型
        ObjectNode p0 = mapper.createObjectNode();
        p0.put("color", "#ffffff");
        p0.set("value", typedValue("percent", 30));
        config.add(p0);
        ObjectNode p1 = mapper.createObjectNode();
        p1.put("color", "#000000");
        p1.set("value", typedValue("percentile", 70));
        config.add(p1);
        ObjectNode p2 = mapper.createObjectNode();
        p2.put("color", "#abcdef");
        ObjectNode v = mapper.createObjectNode();
        v.put("type", "formula");
        v.put("value", "AVERAGE(A1:A10)");
        p2.set("value", v);
        config.add(p2);
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
        JsonNode cfg = rt.get(0).path("rule").path("config");
        assertThat(cfg.size()).isEqualTo(3);
    }

    @Test
    void should_round_trip_dataBar_with_default_color_when_color_missing() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        config.set("min", typedValue("min", 0));
        config.set("max", typedValue("max", 100));
        // 不设 positiveColor / nativeColor → 走默认色分支
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
    }

    @Test
    void should_round_trip_dataBar_with_only_native_color() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        config.set("min", typedValue("num", 0));
        config.set("max", typedValue("num", 100));
        config.put("nativeColor", "#9999ff");
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
    }

    @Test
    void should_use_invalid_color_fallback() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        config.set("min", typedValue("num", 0));
        config.set("max", typedValue("num", 100));
        config.put("positiveColor", "not-a-hex");
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
    }

    @Test
    void should_expose_icon_sentinel_max() {
        // 简单调用以覆盖该静态访问
        assertThat(ConditionalFormattingConverter.iconSentinelMax()).isPositive();
    }

    @Test
    void should_handle_threshold_with_text_value_falling_back_to_zero() throws Exception {
        ArrayNode rules = mapper.createArrayNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.set("ranges", oneRange(0, 0, 9, 0));
        ObjectNode body = mapper.createObjectNode();
        body.put("type", "dataBar");
        ObjectNode config = mapper.createObjectNode();
        ObjectNode minV = mapper.createObjectNode();
        minV.put("type", "num");
        minV.put("value", "abc"); // 非法字符串 → 应走 catch 兜底为 0
        config.set("min", minV);
        config.set("max", typedValue("num", 100));
        body.set("config", config);
        rule.set("rule", body);
        rules.add(rule);
        ArrayNode rt = writeReadBack(rules);
        assertThat(rt.size()).isEqualTo(1);
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
        // 提供 style 触发 applyStyle 路径
        ObjectNode style = mapper.createObjectNode();
        ObjectNode bg = mapper.createObjectNode();
        bg.put("rgb", "#eeeeee");
        style.set("bg", bg);
        body.set("style", style);
        rule.set("rule", body);
        return rule;
    }

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

    private ArrayNode oneRange(int sr, int sc, int er, int ec) {
        ArrayNode arr = mapper.createArrayNode();
        ObjectNode r = mapper.createObjectNode();
        r.put("startRow", sr);
        r.put("startColumn", sc);
        r.put("endRow", er);
        r.put("endColumn", ec);
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
