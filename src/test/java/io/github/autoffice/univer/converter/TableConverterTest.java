package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 表格（Table / ListObject）转换器测试。
 * TableConverter bidirectional tests.
 */
class TableConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_when_sheet_has_no_tables() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("A");
            ObjectNode got = TableConverter.readSheetTables(sh, mapper);
            assertThat(got).isEmpty();
        }
    }

    @Test
    void should_read_table_with_columns_style_and_bounds() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("Name");
            sh.getRow(0).createCell(1).setCellValue("Value");
            sh.createRow(1).createCell(0).setCellValue("foo");
            sh.getRow(1).createCell(1).setCellValue("1");
            sh.createRow(2).createCell(0).setCellValue("bar");
            sh.getRow(2).createCell(1).setCellValue("2");
            XSSFTable table = sh.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));
            table.getCTTable().setRef("A1:B3");
            table.setName("MyTable");
            table.setDisplayName("MyTable");
            table.getCTTable().setHeaderRowCount(1);
            table.getCTTable().getTableColumns().getTableColumnArray(0).setName("Name");
            table.getCTTable().getTableColumns().getTableColumnArray(1).setName("Value");
            table.getCTTable().addNewTableStyleInfo().setName("TableStyleMedium2");
            table.getCTTable().getTableStyleInfo().setShowRowStripes(true);

            ObjectNode got = TableConverter.readSheetTables(sh, mapper);
            assertThat(got.size()).isEqualTo(1);
            JsonNode t = got.fields().next().getValue();
            assertThat(t.path("name").asText()).isEqualTo("MyTable");
            assertThat(t.path("displayName").asText()).isEqualTo("MyTable");
            assertThat(t.path("ref").asText()).isEqualTo("A1:B3");
            assertThat(t.path("startRow").asInt()).isEqualTo(0);
            assertThat(t.path("endRow").asInt()).isEqualTo(2);
            assertThat(t.path("startColumn").asInt()).isEqualTo(0);
            assertThat(t.path("endColumn").asInt()).isEqualTo(1);
            assertThat(t.path("headerRowCount").asInt()).isEqualTo(1);

            JsonNode cols = t.path("columns");
            assertThat(cols.isArray()).isTrue();
            assertThat(cols.size()).isEqualTo(2);
            assertThat(cols.get(0).path("name").asText()).isEqualTo("Name");
            assertThat(cols.get(1).path("name").asText()).isEqualTo("Value");

            JsonNode style = t.path("style");
            assertThat(style.path("name").asText()).isEqualTo("TableStyleMedium2");
            assertThat(style.path("showRowStripes").asBoolean()).isTrue();
        }
    }

    @Test
    void should_write_table_with_ref_columns_and_style() throws Exception {
        ObjectNode rules = mapper.createObjectNode();
        ObjectNode t = mapper.createObjectNode();
        t.put("name", "Tbl1");
        t.put("displayName", "Tbl1");
        t.put("ref", "A1:B3");
        t.put("startRow", 0);
        t.put("endRow", 2);
        t.put("startColumn", 0);
        t.put("endColumn", 1);
        t.put("headerRowCount", 1);
        ArrayNode cols = mapper.createArrayNode();
        cols.add(mapper.createObjectNode().put("id", 1).put("name", "Alpha"));
        cols.add(mapper.createObjectNode().put("id", 2).put("name", "Beta"));
        t.set("columns", cols);
        ObjectNode style = mapper.createObjectNode();
        style.put("name", "TableStyleLight1");
        style.put("showRowStripes", true);
        style.put("showColumnStripes", false);
        t.set("style", style);
        rules.set("1", t);

        ObjectNode got = writeReadBack(rules);
        assertThat(got.size()).isEqualTo(1);
        JsonNode one = got.fields().next().getValue();
        assertThat(one.path("name").asText()).isEqualTo("Tbl1");
        assertThat(one.path("ref").asText()).isEqualTo("A1:B3");
        JsonNode colsBack = one.path("columns");
        assertThat(colsBack.size()).isEqualTo(2);
        assertThat(colsBack.get(0).path("name").asText()).isEqualTo("Alpha");
        assertThat(colsBack.get(1).path("name").asText()).isEqualTo("Beta");
        assertThat(one.path("style").path("name").asText()).isEqualTo("TableStyleLight1");
        assertThat(one.path("style").path("showRowStripes").asBoolean()).isTrue();
    }

    @Test
    void should_skip_duplicate_tables_when_same_ref_already_exists() throws Exception {
        ObjectNode tables = mapper.createObjectNode();
        ObjectNode t = mapper.createObjectNode();
        t.put("name", "Dup");
        t.put("ref", "A1:B2");
        t.put("startRow", 0);
        t.put("endRow", 1);
        t.put("startColumn", 0);
        t.put("endColumn", 1);
        tables.set("1", t);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("a");
            sh.getRow(0).createCell(1).setCellValue("b");
            sh.createRow(1).createCell(0).setCellValue("1");
            sh.getRow(1).createCell(1).setCellValue("2");
            // pre-existing table
            XSSFTable existing = sh.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
            existing.getCTTable().setRef("A1:B2");
            existing.setName("Dup");
            existing.setDisplayName("Dup");

            TableConverter.writeSheetTables(sh, tables);
            List<XSSFTable> after = sh.getTables();
            assertThat(after).hasSize(1);
        }
    }

    private ObjectNode writeReadBack(ObjectNode tables) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            // Tables require populated headers or POI will still accept but Excel expects data cells.
            sh.createRow(0).createCell(0).setCellValue("Alpha");
            sh.getRow(0).createCell(1).setCellValue("Beta");
            sh.createRow(1).createCell(0).setCellValue("x");
            sh.getRow(1).createCell(1).setCellValue("y");
            sh.createRow(2).createCell(0).setCellValue("p");
            sh.getRow(2).createCell(1).setCellValue("q");
            TableConverter.writeSheetTables(sh, tables);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return TableConverter.readSheetTables(wb2.getSheetAt(0), mapper);
        }
    }
}
