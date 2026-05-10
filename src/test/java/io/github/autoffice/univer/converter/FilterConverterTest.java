package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 自动筛选转换器测试。
 * FilterConverter bidirectional tests.
 */
class FilterConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_when_sheet_has_no_autoFilter() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("A");
            ObjectNode got = FilterConverter.readSheetFilter(sh, mapper);
            assertThat(got).isEmpty();
        }
    }

    @Test
    void should_read_ref_and_numeric_bounds_when_sheet_has_autoFilter() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("H1");
            sh.getRow(0).createCell(1).setCellValue("H2");
            sh.createRow(1).createCell(0).setCellValue("1");
            sh.getRow(1).createCell(1).setCellValue("2");
            sh.setAutoFilter(CellRangeAddress.valueOf("A1:B2"));

            ObjectNode got = FilterConverter.readSheetFilter(sh, mapper);
            assertThat(got.path("ref").asText()).isEqualTo("A1:B2");
            assertThat(got.path("startRow").asInt()).isEqualTo(0);
            assertThat(got.path("endRow").asInt()).isEqualTo(1);
            assertThat(got.path("startColumn").asInt()).isEqualTo(0);
            assertThat(got.path("endColumn").asInt()).isEqualTo(1);
            assertThat(got.has("rawXml")).isTrue();
        }
    }

    @Test
    void should_write_filter_from_ref_only_when_no_rawXml() throws Exception {
        byte[] bytes = writeThenSerialize(null, "A1:C3", 0, 2, 0, 2);
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet sh = wb2.getSheetAt(0);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
            assertThat(sh.getCTWorksheet().getAutoFilter().getRef()).isEqualTo("A1:C3");
        }
    }

    @Test
    void should_roundtrip_autoFilter_with_rawXml_preservation() throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            for (int r = 0; r < 4; r++) {
                sh.createRow(r).createCell(0).setCellValue(r == 0 ? "Col" : String.valueOf(r));
            }
            sh.setAutoFilter(CellRangeAddress.valueOf("A1:A4"));
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }

        ObjectNode firstRead;
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            firstRead = FilterConverter.readSheetFilter(wb2.getSheetAt(0), mapper);
        }
        assertThat(firstRead.path("ref").asText()).isEqualTo("A1:A4");

        byte[] bytes2;
        try (XSSFWorkbook wb3 = new XSSFWorkbook()) {
            XSSFSheet sh = wb3.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("Col");
            FilterConverter.writeSheetFilter(sh, firstRead);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb3.write(out);
                bytes2 = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb4 = new XSSFWorkbook(new ByteArrayInputStream(bytes2))) {
            XSSFSheet sh = wb4.getSheetAt(0);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
            assertThat(sh.getCTWorksheet().getAutoFilter().getRef()).isEqualTo("A1:A4");
        }
    }

    @Test
    void should_skip_write_when_payload_missing_range() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("A");
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    private byte[] writeThenSerialize(String rawXml, String ref,
                                      int startRow, int endRow, int startCol, int endCol) throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        if (ref != null) {
            payload.put("ref", ref);
        }
        payload.put("startRow", startRow);
        payload.put("endRow", endRow);
        payload.put("startColumn", startCol);
        payload.put("endColumn", endCol);
        if (rawXml != null) {
            payload.put("rawXml", rawXml);
        }
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            for (int r = startRow; r <= endRow; r++) {
                sh.createRow(r).createCell(startCol).setCellValue("v" + r);
            }
            FilterConverter.writeSheetFilter(sh, payload);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                return out.toByteArray();
            }
        }
    }
}
