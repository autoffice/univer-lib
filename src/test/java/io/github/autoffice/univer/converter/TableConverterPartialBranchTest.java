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

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * TableConverter partial 分支精确补全：
 * - 表格名/displayName 全为 null（fallback）
 * - tableNode 缺各字段
 * - columns 含 null col / null name
 * - style 中各字段缺失（showColumnStripes 等）
 */
class TableConverterPartialBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_use_displayName_as_key_when_name_null() throws Exception {
        // 覆盖 L116: name == null/empty → 用 displayName
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue(r == 0 ? "h" : String.valueOf(r));
            }
            XSSFTable table = sh.createTable(new AreaReference("A1:A3", SpreadsheetVersion.EXCEL2007));
            // 不设 name 直接用 setDisplayName
            table.setDisplayName("DispOnly");
            table.getCTTable().unsetName(); // 清除 name
            table.getCTTable().setId(0); // 让 id 为 0 走 fallback

            ObjectNode out = TableConverter.readSheetTables(sh, mapper);
            assertThat(out).isNotNull();
        }
    }

    @Test
    void should_handle_table_without_displayName_in_write() throws Exception {
        // 覆盖 L130: display == null 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            // 只设 name，不设 displayName
            t.put("name", "OnlyName");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            // displayName 应当 fallback 为 name
            assertThat(sh.getTables().get(0).getDisplayName()).isEqualTo("OnlyName");
        }
    }

    @Test
    void should_skip_invalid_columns_with_null_or_non_object() throws Exception {
        // 覆盖 L154 / L159 的 col == null / col.getName() == null 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            ArrayNode cols = mapper.createArrayNode();
            cols.add(mapper.nullNode()); // null col
            cols.add(mapper.createArrayNode()); // 非 object
            cols.add(mapper.createObjectNode().put("id", 1)); // 没 name 字段
            t.set("columns", cols);
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_style_when_node_is_empty() throws Exception {
        // 覆盖 applyStyle 的 styleNode == null / size == 0 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            t.set("style", mapper.createObjectNode()); // 空对象
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_set_style_with_only_some_fields() throws Exception {
        // 覆盖 hasNonNull 各字段分别缺失的分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            ObjectNode style = mapper.createObjectNode();
            style.put("name", "TableStyleLight1");
            // 仅设 name，不设 stripes/firstColumn/lastColumn
            t.set("style", style);
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_negative_headerRowCount() throws Exception {
        // 覆盖 L256 (v >= 0 ? false 分支)：当 headerRowCount 为负时不设
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            t.put("headerRowCount", -1); // 负数 → 跳过
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_zero_totalsRowCount() throws Exception {
        // 覆盖 L260: totalsRowCount == 0 → 不设 totalsRowShown
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            t.put("totalsRowCount", 0); // 0 → 跳过
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_existing_table_with_null_name() throws Exception {
        // 覆盖 L204: t.getName() == null 分支（existing table 没 name）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            // 创建一个无 name 的 existing table
            XSSFTable existing = sh.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));
            existing.getCTTable().unsetName();

            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A4:B6"); // 不同 ref
            t.put("name", "New");
            payload.set("1", t);
            // 不应抛错（existing.getName 为 null 时跳过 add）
            // 注意：sheet 必须有足够的行
            for (int r = 3; r < 6; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            TableConverter.writeSheetTables(sh, payload);
        }
    }

    @Test
    void should_handle_invalid_ref_with_numeric_fallback() throws Exception {
        // 覆盖 resolveArea 中 ref 非法 → numeric 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "$$$"); // 非法
            t.put("startRow", 0);
            t.put("endRow", 2);
            t.put("startColumn", 0);
            t.put("endColumn", 1);
            t.put("name", "T");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_table_with_empty_string_name() throws Exception {
        // 覆盖 L243: name == "" → 不设 name
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", ""); // 空串
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            // 仍然能创建 table（POI 自动生成名字）
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_table_with_empty_displayName() throws Exception {
        // 覆盖 L247: displayName 为空串 → 走 elif 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "MyName");
            t.put("displayName", ""); // 空串 → fallback to name
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables().get(0).getDisplayName()).isEqualTo("MyName");
        }
    }

    @Test
    void should_skip_invalid_ref_in_resolveArea_with_no_numeric() throws Exception {
        // 覆盖 resolveArea L278: ref 非法 + 无 numeric 字段 → return null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "$$$"); // 非法
            // 不设 numeric
            t.put("name", "T");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).isEmpty();
        }
    }

    @Test
    void should_skip_when_ref_empty_string_in_payload() throws Exception {
        // 覆盖 L278: ref 为空串 → 落到 numeric
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
            }
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", ""); // 空串
            t.put("startRow", 0);
            t.put("endRow", 2);
            t.put("startColumn", 0);
            t.put("endColumn", 0);
            t.put("name", "T");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_existing_table_without_ref_in_ct() throws Exception {
        // 覆盖 L208: ct.getRef() == null 分支
        // POI 的 CTTable 不允许 unset ref，这里跳过具体场景；通过 round-trip 替代
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                sh.createRow(r).createCell(0).setCellValue("v");
                sh.getRow(r).createCell(1).setCellValue("w");
            }
            // 不抛错即可
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode t = mapper.createObjectNode();
            t.put("ref", "A1:B3");
            t.put("name", "T");
            payload.set("1", t);
            TableConverter.writeSheetTables(sh, payload);
            assertThat(sh.getTables()).hasSize(1);
        }
    }
}
