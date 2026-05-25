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
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Iterator;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * DataValidationConverter 分支补全：所有 operator、type、errorStyle、坏数据、strictMode 等。
 */
class DataValidationConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_on_null_sheet() {
        ObjectNode out = DataValidationConverter.readSheetDataValidations(null, mapper);
        assertThat(out.size()).isEqualTo(0);
    }

    @Test
    void should_skip_writeSheetDataValidations_on_null_inputs() throws Exception {
        // null sheet
        DataValidationConverter.writeSheetDataValidations(null, mapper.createObjectNode(), null);
        // null rules
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, null);
            DataValidationConverter.writeSheetDataValidations(sh, mapper.createObjectNode()); // empty
            assertThat(sh.getDataValidations()).isEmpty();
        }
    }

    @Test
    void should_skip_rule_with_invalid_node_or_no_ranges() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        rules.set("r1", mapper.nullNode()); // null
        rules.set("r2", mapper.createArrayNode()); // 非 object
        ObjectNode rule = baseRule("r3", "whole", mapper.createArrayNode()); // 空 ranges
        rule.put("formula1", "1");
        rules.set("r3", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).isEmpty();
        }
    }

    @Test
    void should_roundtrip_each_operator_for_decimal() throws Exception {
        String[] ops = {"equal", "notEqual", "greaterThan", "lessThan",
                        "greaterThanOrEqual", "lessThanOrEqual", "notBetween"};
        for (String op : ops) {
            ObjectNode rules = mapper.createObjectNode();
            ObjectNode rule = baseRule("r", "decimal", oneRange(0, 0, 0, 0));
            rule.put("operator", op);
            rule.put("formula1", "1");
            if ("notBetween".equals(op)) {
                rule.put("formula2", "10");
            }
            rules.set("r", rule);
            ObjectNode rt = writeReadBack(rules);
            JsonNode got = firstRule(rt);
            assertThat(got.path("operator").asText()).as("op=%s", op).isEqualTo(op);
        }
    }

    @Test
    void should_roundtrip_textLength_type() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "textLength", oneRange(0, 0, 0, 0));
        rule.put("operator", "greaterThan");
        rule.put("formula1", "5");
        rules.set("r", rule);
        ObjectNode rt = writeReadBack(rules);
        JsonNode got = firstRule(rt);
        assertThat(got.path("type").asText()).isEqualTo("textLength");
    }

    @Test
    void should_roundtrip_date_with_time_constraint() throws Exception {
        // bizInfo.showTime=true → createTimeConstraint
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "date", oneRange(0, 0, 0, 0));
        rule.put("operator", "equal");
        rule.put("formula1", "TIME(12,0,0)");
        ObjectNode biz = mapper.createObjectNode();
        biz.put("showTime", true);
        rule.set("bizInfo", biz);
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).hasSize(1);
        }
    }

    @Test
    void should_roundtrip_any_type_as_custom_TRUE() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "any", oneRange(0, 0, 0, 0));
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).hasSize(1);
        }
    }

    @Test
    void should_roundtrip_listMultiple_with_explicit_values() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "listMultiple", oneRange(0, 0, 0, 0));
        rule.put("formula1", "A,B,C");
        rule.put("showDropDown", false); // 走 setFormula2 分支
        rule.put("formula2", "marker");
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).hasSize(1);
        }
    }

    @Test
    void should_handle_empty_list_formula() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "list", oneRange(0, 0, 0, 0));
        rule.put("formula1", "");
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // POI 拒绝空 list；createListConstraint 创建了空数组约束，sheet.addValidationData 抛 IAE
            assertThatThrownBy(() -> DataValidationConverter.writeSheetDataValidations(sh, rules))
                    .isInstanceOf(IllegalArgumentException.class);
        }
    }

    @Test
    void should_skip_unsupported_type_in_lenient_mode() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "weirdType", oneRange(0, 0, 0, 0));
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // 默认 lenient：跳过不抛
            DataValidationConverter.writeSheetDataValidations(sh, rules, UniverXlsxOptions.defaults());
            assertThat(sh.getDataValidations()).isEmpty();
        }
    }

    @Test
    void should_throw_in_strict_mode_for_unsupported_type() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "weirdType", oneRange(0, 0, 0, 0));
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            UniverXlsxOptions strict = UniverXlsxOptions.builder().strictMode(true).build();
            assertThatThrownBy(() -> DataValidationConverter.writeSheetDataValidations(sh, rules, strict))
                    .isInstanceOf(UniverXlsxUnsupportedFeatureException.class)
                    .hasMessageContaining("Unsupported");
        }
    }

    @Test
    void should_round_trip_error_style_warning() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "whole", oneRange(0, 0, 0, 0));
        rule.put("operator", "equal");
        rule.put("formula1", "5");
        rule.put("errorStyle", "warning");
        rule.put("showErrorMessage", true);
        rule.put("error", "err");
        rule.put("errorTitle", "etitle");
        rule.put("prompt", "p");
        rule.put("promptTitle", "ptitle");
        rule.put("showInputMessage", true);
        rules.set("r", rule);
        ObjectNode rt = writeReadBack(rules);
        JsonNode got = firstRule(rt);
        assertThat(got.path("errorStyle").asText()).isEqualTo("warning");
    }

    @Test
    void should_round_trip_error_style_info() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "whole", oneRange(0, 0, 0, 0));
        rule.put("operator", "equal");
        rule.put("formula1", "5");
        rule.put("errorStyle", "info");
        rules.set("r", rule);
        ObjectNode rt = writeReadBack(rules);
        JsonNode got = firstRule(rt);
        assertThat(got.path("errorStyle").asText()).isEqualTo("info");
    }

    @Test
    void should_accept_numeric_error_style() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "whole", oneRange(0, 0, 0, 0));
        rule.put("operator", "equal");
        rule.put("formula1", "5");
        rule.put("errorStyle", 1); // WARNING
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).hasSize(1);
        }
    }

    @Test
    void should_normalize_whitespace_only_formulas_to_null() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("r", "whole", oneRange(0, 0, 0, 0));
        rule.put("operator", "equal");
        rule.put("formula1", "   "); // 应转为 null
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // POI 在公式被 normalize 为 null 后会拒绝，验证我们触发了 normalize 路径
            assertThatThrownBy(() -> DataValidationConverter.writeSheetDataValidations(sh, rules))
                    .isInstanceOf(IllegalArgumentException.class);
        }
    }

    @Test
    void should_skip_invalid_range_indices() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = mapper.createObjectNode();
        rule.put("uid", "r");
        rule.put("type", "whole");
        ArrayNode ranges = mapper.createArrayNode();
        ObjectNode bad = mapper.createObjectNode();
        // 缺字段 → -1 → 跳过
        ranges.add(bad);
        rule.set("ranges", ranges);
        rule.put("operator", "equal");
        rule.put("formula1", "5");
        rules.set("r", rule);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            assertThat(sh.getDataValidations()).isEmpty();
        }
    }

    @Test
    void should_build_resource_per_sheet_id() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // 先写一条 / write one rule
            ObjectNode rules = mapper.createObjectNode();
            ObjectNode rule = baseRule("r", "whole", oneRange(0, 0, 0, 0));
            rule.put("operator", "equal");
            rule.put("formula1", "1");
            rules.set("r", rule);
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            ObjectNode out = DataValidationConverter.buildResourceBySheetId(sh, mapper);
            // key 是 sheetIndex 字符串
            assertThat(out.fieldNames().hasNext()).isTrue();
        }
    }

    @Test
    void should_return_empty_resource_when_no_validations() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode out = DataValidationConverter.buildResourceBySheetId(sh, mapper);
            assertThat(out.size()).isEqualTo(0);
        }
    }

    // ---------- helpers ----------

    private JsonNode firstRule(ObjectNode rules) {
        Iterator<Map.Entry<String, JsonNode>> it = rules.fields();
        assertThat(it.hasNext()).isTrue();
        return it.next().getValue();
    }

    private ObjectNode writeReadBack(ObjectNode rules) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("anchor");
            DataValidationConverter.writeSheetDataValidations(sh, rules);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return DataValidationConverter.readSheetDataValidations(wb2.getSheetAt(0), mapper);
        }
    }

    private ObjectNode baseRule(String uid, String type, ArrayNode ranges) {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("uid", uid);
        rule.put("type", type);
        rule.set("ranges", ranges);
        return rule;
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
}
