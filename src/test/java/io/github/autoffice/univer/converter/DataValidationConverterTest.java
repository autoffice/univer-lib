package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
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

/**
 * 数据验证双向转换测试。
 * Bidirectional data validation conversion tests.
 */
class DataValidationConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_roundtrip_whole_number_between_when_write_and_read_sheet_validations() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("rule1", "whole", oneRange(0, 0, 9, 0));
        rule.put("operator", "between");
        rule.put("formula1", "1");
        rule.put("formula2", "10");
        rule.put("allowBlank", true);
        rule.put("showErrorMessage", true);
        rule.put("error", "Please enter 1-10");
        rules.set("rule1", rule);

        ObjectNode roundTrip = writeReadBack(rules, null);
        JsonNode got = firstRule(roundTrip);
        assertThat(got.path("type").asText()).isEqualTo("whole");
        assertThat(got.path("operator").asText()).isEqualTo("between");
        assertThat(got.path("formula1").asText()).isEqualTo("1");
        assertThat(got.path("formula2").asText()).isEqualTo("10");
        assertThat(got.path("showErrorMessage").asBoolean()).isTrue();
        assertThat(got.path("error").asText()).isEqualTo("Please enter 1-10");
    }

    @Test
    void should_roundtrip_explicit_list_when_write_and_read_sheet_validations() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("rule1", "list", oneRange(0, 0, 1, 0));
        rule.put("formula1", "Yes,No");
        rule.put("showDropDown", true);
        rules.set("rule1", rule);

        ObjectNode roundTrip = writeReadBack(rules, null);
        JsonNode got = firstRule(roundTrip);
        assertThat(got.path("type").asText()).isEqualTo("list");
        assertThat(got.path("formula1").asText()).isEqualTo("Yes,No");
        assertThat(got.path("showDropDown").asBoolean()).isTrue();
    }

    @Test
    void should_roundtrip_formula_list_when_write_and_read_sheet_validations() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("rule1", "list", oneRange(0, 0, 1, 0));
        rule.put("formula1", "=B1:B2");
        rule.put("showDropDown", true);
        rules.set("rule1", rule);

        ObjectNode roundTrip = writeReadBack(rules, workbookWithListSource());
        JsonNode got = firstRule(roundTrip);
        assertThat(got.path("type").asText()).isEqualTo("list");
        assertThat(got.path("formula1").asText()).isEqualTo("=B1:B2");
    }

    @Test
    void should_roundtrip_custom_formula_when_write_and_read_sheet_validations() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("rule1", "custom", oneRange(0, 0, 9, 0));
        rule.put("formula1", "=COUNTIF($B$1:$B$3,A1)>0");
        rules.set("rule1", rule);

        ObjectNode roundTrip = writeReadBack(rules, workbookWithListSource());
        JsonNode got = firstRule(roundTrip);
        assertThat(got.path("type").asText()).isEqualTo("custom");
        assertThat(got.path("formula1").asText()).isEqualTo("COUNTIF($B$1:$B$3,A1)>0");
    }

    @Test
    void should_roundtrip_date_validation_when_write_and_read_sheet_validations() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode rule = baseRule("rule1", "date", oneRange(0, 0, 9, 0));
        rule.put("operator", "between");
        rule.put("formula1", "DATE(2024,1,1)");
        rule.put("formula2", "DATE(2024,12,31)");
        rules.set("rule1", rule);

        ObjectNode roundTrip = writeReadBack(rules, null);
        JsonNode got = firstRule(roundTrip);
        assertThat(got.path("type").asText()).isEqualTo("date");
        assertThat(got.path("operator").asText()).isEqualTo("between");
        assertThat(got.path("formula1").asText()).isEqualTo("DATE(2024,1,1)");
        assertThat(got.path("formula2").asText()).isEqualTo("DATE(2024,12,31)");
    }

    @Test
    void should_write_native_poi_validations_from_univer_rules() throws Exception {
        try (XSSFWorkbook wb = workbookWithListSource()) {
            XSSFSheet sheet = wb.getSheetAt(0);
            ObjectNode rules = mapper.createObjectNode();
            ObjectNode rule = baseRule("rule1", "list", oneRange(0, 0, 1, 0));
            rule.put("formula1", "=B1:B2");
            rule.put("showDropDown", true);
            rules.set("rule1", rule);

            DataValidationConverter.writeSheetDataValidations(sheet, rules);

            assertThat(sheet.getDataValidations()).hasSize(1);
            DataValidation validation = sheet.getDataValidations().get(0);
            assertThat(validation.getValidationConstraint().getFormula1()).isEqualTo("B1:B2");
        }
    }

    private JsonNode firstRule(ObjectNode rules) {
        Iterator<Map.Entry<String, JsonNode>> it = rules.fields();
        assertThat(it.hasNext()).as("roundTrip should contain at least one rule").isTrue();
        return it.next().getValue();
    }

    private ObjectNode writeReadBack(ObjectNode rules, XSSFWorkbook seedWorkbook) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = seedWorkbook == null ? new XSSFWorkbook() : seedWorkbook) {
            XSSFSheet sh = wb.getNumberOfSheets() == 0 ? wb.createSheet("S1") : wb.getSheetAt(0);
            if (sh.getPhysicalNumberOfRows() == 0) {
                sh.createRow(0).createCell(0).setCellValue("anchor");
            }
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

    private XSSFWorkbook workbookWithListSource() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("S1");
        sheet.createRow(0).createCell(1).setCellValue("Yes");
        sheet.getRow(0).createCell(0).setCellValue("seed");
        sheet.createRow(1).createCell(1).setCellValue("No");
        sheet.getRow(1).createCell(0).setCellValue("seed");
        return wb;
    }

    private ObjectNode baseRule(String uid, String type, ArrayNode ranges) {
        ObjectNode rule = mapper.createObjectNode();
        rule.put("uid", uid);
        rule.put("type", type);
        rule.set("ranges", ranges);
        return rule;
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
}
