/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */
package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * TableConverter 边角分支：null sheet/payload、numeric-only ref、坏 ref、totalsRowCount、空 columns、style 兜底。
 */
class TableConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_on_null_sheet() {
        ObjectNode out = TableConverter.readSheetTables(null, mapper);
        assertThat(out).isEmpty();
    }

    @Test
    void should_skip_write_when_payload_invalid() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("a");
            TableConverter.writeSheetTables(sh, null);
            TableConverter.writeSheetTables(sh, mapper.createObjectNode());
            TableConverter.writeSheetTables(sh, mapper.createArrayNode());
            TableConverter.writeSheetTables(null, mapper.createObjectNode());
            assertThat(sh.getTables()).isEmpty();
        }
    }

    @Test
    void should_skip_table_with_invalid_node() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.set("t1", mapper.nullNode());
            payload.set("t2", mapper.createArrayNode()); // 非 object
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).isEmpty();
        }
    }

    @Test
    void should_use_numeric_range_when_ref_missing() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("Alpha");
            sh.getRow(0).createCell(1).setCellValue("Beta");
            sh.createRow(1).createCell(0).setCellValue("x");
            sh.getRow(1).createCell(1).setCellValue("y");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            // 没 ref，只有 numeric 字段
            t.put("startRow", 0);
            t.put("endRow", 1);
            t.put("startColumn", 0);
            t.put("endColumn", 1);
            t.put("name", "NumOnly");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_when_neither_ref_nor_numeric_fields() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("name", "NoRange");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).isEmpty();
        }
    }

    @Test
    void should_set_totalsRowCount_when_positive() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 4; r++) {
                sh.createRow(r).createCell(0).setCellValue(r == 0 ? "Header" : String.valueOf(r));
                sh.getRow(r).createCell(1).setCellValue(r == 0 ? "B" : String.valueOf(r));
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B4");
            t.put("name", "WithTotals");
            t.put("totalsRowCount", 1);
            t.put("headerRowCount", 1);
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
            assertThat(sh.getTables().get(0).getCTTable().getTotalsRowCount()).isEqualTo(1);
        }
    }

    @Test
    void should_skip_table_with_invalid_ref_and_no_numeric_range() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "$$$");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).isEmpty();
        }
    }

    @Test
    void should_skip_invalid_columns_entries() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue(r == 0 ? "h1" : String.valueOf(r));
                sh.getRow(r).createCell(1).setCellValue(r == 0 ? "h2" : String.valueOf(r));
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "BadCols");
            // 列数组中混杂非对象
            com.fasterxml.jackson.databind.node.ArrayNode cols = mapper.createArrayNode();
            cols.add(mapper.nullNode());
            cols.add(mapper.createObjectNode().put("name", "Real"));
            t.set("columns", cols);
            // style 空对象 → 跳过 applyStyle
            t.set("style", mapper.createObjectNode());
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_use_displayName_fallback_when_name_present() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "OnlyName");
            // 不设 displayName，触发 fallback 设为 name
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables().get(0).getDisplayName()).isEqualTo("OnlyName");
        }
    }

    @Test
    void should_read_table_without_columns_section() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue(r == 0 ? "h" : String.valueOf(r));
            }
            // 创建表（不写 name），让 POI 自动生成默认 name；断言读路径仍能正常返回 1 条记录
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:A3");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            ObjectNode out = TableConverter.readSheetTables(sh, mapper);
            assertThat(out.size()).isEqualTo(1);
            JsonNode one = out.fields().next().getValue();
            assertThat(one.path("ref").asText()).isEqualTo("A1:A3");
        }
    }

    @Test
    void should_skip_table_when_name_already_exists() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode first = mapper.createObjectNode();
            ObjectNode t1 = mapper.createObjectNode();
            t1.put("ref", "A1:B3");
            t1.put("name", "Same");
            first.set("1", t1);
            TableConverter.writeSheetTables(sh, first);

            // 同名再写一次 → 应被跳过
            ObjectNode second = mapper.createObjectNode();
            ObjectNode t2 = mapper.createObjectNode();
            t2.put("ref", "A1:B3");
            t2.put("name", "Same");
            second.set("2", t2);
            TableConverter.writeSheetTables(sh, second);
            assertThat(sh.getTables()).hasSize(1);
        }
    }
}
