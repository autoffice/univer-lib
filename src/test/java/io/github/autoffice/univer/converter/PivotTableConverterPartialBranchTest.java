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
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * PivotTableConverter partial 分支精确补全：
 * - readPivotTable 中各 null 字段 (unitId, subUnitId, def, name, ref 等)
 * - writeSheetPivotTables 多个 valid pivot
 * - applyOptions 中各项缺失
 * - subtotalToUniverName 各种 enum 值
 */
class PivotTableConverterPartialBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_read_pivot_with_null_unitId_and_subUnitId() throws Exception {
        // 覆盖 L106, L109, L156, L215, L216: unitId/subUnitId 为 null 的三元运算符
        byte[] bytes = workbookWithPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            // unitId=null, subUnitId=null
            ArrayNode out = PivotTableConverter.readSheetPivotTables(wb, pivotSheet, mapper, null, null);
            assertThat(out.size()).isEqualTo(1);
            assertThat(out.get(0).path("sourceRangeInfo").path("unitId").asText()).isEmpty();
            assertThat(out.get(0).path("targetCellInfo").path("subUnitId").asText()).isEqualTo("Pivot");
        }
    }

    @Test
    void should_handle_each_subtotal_function_in_read() throws Exception {
        // 通过创建带不同 subtotal 的 pivot table 覆盖 subtotalToUniverName 各分支
        byte[] bytes = workbookWithMultiSubtotalPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode out = PivotTableConverter.readSheetPivotTables(wb, pivotSheet, mapper, "u", "Pivot");
            assertThat(out.size()).isEqualTo(1);
        }
    }

    @Test
    void should_handle_pivot_with_pivot_payload_array_partial_invalid() throws Exception {
        // 覆盖 writePivotTable 中各分支
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            // 完整 pivot
            payload.add(buildValidPivotNode());
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_apply_pivot_options_with_grand_totals() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            ObjectNode opts = mapper.createObjectNode();
            opts.put("showRowGrandTotal", true);  // 默认 true，覆盖另一分支
            opts.put("showColumnGrandTotal", true);
            opts.put("compact", true);
            opts.put("outline", false);
            opts.put("outlineData", false);
            opts.put("compactData", true);
            opts.put("multipleFieldFilters", false);
            opts.put("dataCaption", "DataValues");
            opts.put("dataOnRows", false);
            p.set("options", opts);
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_apply_pivot_with_no_options_node() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            // 不设 options
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_apply_pivot_options_with_array_node_skipping() throws Exception {
        // options 是 array，不是 object → applyOptions 应跳过
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            p.set("options", mapper.createArrayNode()); // 非 object
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_pivot_with_filters() throws Exception {
        // 覆盖 applyFieldConfig 中的 filters 分支
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            ObjectNode fc = (ObjectNode) p.get("fieldsConfig");
            ArrayNode filters = mapper.createArrayNode();
            filters.add(mapper.createObjectNode().put("sourceIndex", 0));
            filters.add(mapper.createObjectNode().put("sourceIndex", -1)); // 跳过
            fc.set("filters", filters);
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_pivot_with_negative_source_index_in_rows() throws Exception {
        // 覆盖 applyFieldConfig 中 sourceIndex < 0 的跳过分支
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            ObjectNode fc = (ObjectNode) p.get("fieldsConfig");
            ArrayNode rows = (ArrayNode) fc.get("rows");
            // 添加一个非法 sourceIndex
            rows.add(mapper.createObjectNode().put("sourceIndex", -1));
            ArrayNode columns = (ArrayNode) fc.get("columns");
            columns.add(mapper.createObjectNode().put("sourceIndex", -1));
            ArrayNode values = (ArrayNode) fc.get("values");
            values.add(mapper.createObjectNode().put("sourceIndex", -1));
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_pivot_resolveSourceSheet_with_subUnitId_not_in_map() throws Exception {
        // 覆盖 resolveSourceSheet 中 sheetIdToSheet 找不到 subUnitId 的分支 → fallback to sheetName
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            ObjectNode src = (ObjectNode) p.get("sourceRangeInfo");
            src.put("subUnitId", "ghost"); // map 里没有这个
            // 但 sheetName 仍是 "Source" → 应 fallback
            payload.add(p);
            Map<String, XSSFSheet> map = new LinkedHashMap<>(); // 空 map
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_skip_pivot_resolveSourceSheet_with_null_sheetIdMap() throws Exception {
        // 覆盖 resolveSourceSheet 中 sheetIdToSheet == null 分支
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = buildValidPivotNode();
            ObjectNode src = (ObjectNode) p.get("sourceRangeInfo");
            src.put("subUnitId", "Source"); // 通过 sheetName fallback
            payload.add(p);
            // 不传 map (null)
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, null);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_read_pivot_with_name_dataCaption_and_subtotalCaption() throws Exception {
        // 通过对 CTPivotTableDefinition 直接设置 name / dataCaption 等，覆盖
        // L113, L124, L333 等多分支 partial
        byte[] bytes = workbookWithPivotAndCustomDef();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode out = PivotTableConverter.readSheetPivotTables(wb, pivotSheet, mapper, "u", "Pivot");
            assertThat(out.size()).isEqualTo(1);
            // name 应被读取
            assertThat(out.get(0).path("name").asText()).isEqualTo("MyPivot");
            // dataCaption 也应读取
            assertThat(out.get(0).path("options").path("dataCaption").asText()).isEqualTo("Vals");
        }
    }

    @Test
    void should_read_pivot_with_page_fields() throws Exception {
        // 覆盖 addPageFields 循环 L351 + 字段 name 不为 null 分支 L358
        byte[] bytes = workbookWithPivotAndPageFilter();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode out = PivotTableConverter.readSheetPivotTables(wb, pivotSheet, mapper, "u", "Pivot");
            assertThat(out.size()).isEqualTo(1);
            // filters 数组应非空
            assertThat(out.get(0).path("fieldsConfig").path("filters").size()).isPositive();
        }
    }

    private byte[] workbookWithPivotAndPageFilter() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Quarter");
            h.createCell(2).setCellValue("Amount");
            for (int i = 1; i <= 3; i++) {
                XSSFRow r = src.createRow(i);
                r.createCell(0).setCellValue("R" + i);
                r.createCell(1).setCellValue("Q" + i);
                r.createCell(2).setCellValue(i * 100);
            }
            XSSFSheet ps = wb.createSheet("Pivot");
            XSSFPivotTable pivot = ps.createPivotTable(
                    new AreaReference("A1:C4", SpreadsheetVersion.EXCEL2007),
                    new CellReference(0, 0), src);
            pivot.addReportFilter(0); // page filter
            pivot.addRowLabel(1);
            pivot.addColumnLabel(DataConsolidateFunction.SUM, 2);
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return bos.toByteArray();
        }
    }

    private byte[] workbookWithPivotAndCustomDef() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Amount");
            for (int i = 1; i <= 2; i++) {
                XSSFRow r = src.createRow(i);
                r.createCell(0).setCellValue("R" + i);
                r.createCell(1).setCellValue(i * 100);
            }
            XSSFSheet ps = wb.createSheet("Pivot");
            org.apache.poi.xssf.usermodel.XSSFPivotTable pivot = ps.createPivotTable(
                    new org.apache.poi.ss.util.AreaReference("A1:B3", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007),
                    new org.apache.poi.ss.util.CellReference(0, 0), src);
            pivot.addRowLabel(0);
            pivot.addColumnLabel(org.apache.poi.ss.usermodel.DataConsolidateFunction.SUM, 1);
            // 设置 name 和 dataCaption
            pivot.getCTPivotTableDefinition().setName("MyPivot");
            pivot.getCTPivotTableDefinition().setDataCaption("Vals");
            // 设置 subtotalCaption 在第 0 个 pivotField
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields fields =
                    pivot.getCTPivotTableDefinition().getPivotFields();
            if (fields != null && fields.sizeOfPivotFieldArray() > 0) {
                fields.getPivotFieldArray(0).setSubtotalCaption("MySub");
            }
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return bos.toByteArray();
        }
    }

    private ObjectNode buildValidPivotNode() {
        ObjectNode p = mapper.createObjectNode();
        p.put("pivotTableId", "p1");
        ObjectNode src = mapper.createObjectNode();
        src.put("sheetName", "Source");
        ObjectNode range = mapper.createObjectNode();
        range.put("startRow", 0);
        range.put("startColumn", 0);
        range.put("endRow", 2);
        range.put("endColumn", 2);
        src.set("range", range);
        p.set("sourceRangeInfo", src);
        ObjectNode target = mapper.createObjectNode();
        target.put("sheetName", "Pivot");
        target.put("row", 0);
        target.put("column", 0);
        p.set("targetCellInfo", target);
        ObjectNode fc = mapper.createObjectNode();
        ArrayNode rows = mapper.createArrayNode();
        rows.add(mapper.createObjectNode().put("sourceIndex", 0));
        ArrayNode cols = mapper.createArrayNode();
        cols.add(mapper.createObjectNode().put("sourceIndex", 1));
        ArrayNode values = mapper.createArrayNode();
        values.add(mapper.createObjectNode().put("sourceIndex", 2).put("subtotal", "sum"));
        fc.set("rows", rows);
        fc.set("columns", cols);
        fc.set("filters", mapper.createArrayNode());
        fc.set("values", values);
        p.set("fieldsConfig", fc);
        return p;
    }

    private byte[] workbookWithPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Quarter");
            h.createCell(2).setCellValue("Amount");
            for (int i = 1; i <= 2; i++) {
                XSSFRow r = src.createRow(i);
                r.createCell(0).setCellValue(i == 1 ? "E" : "W");
                r.createCell(1).setCellValue("Q" + i);
                r.createCell(2).setCellValue(i * 100);
            }
            XSSFSheet ps = wb.createSheet("Pivot");
            XSSFPivotTable pivot = ps.createPivotTable(
                    new AreaReference("A1:C3", SpreadsheetVersion.EXCEL2007),
                    new CellReference(0, 0), src);
            pivot.addRowLabel(0);
            pivot.addColLabel(1);
            pivot.addColumnLabel(DataConsolidateFunction.SUM, 2);
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return bos.toByteArray();
        }
    }

    private byte[] workbookWithMultiSubtotalPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Amount");
            for (int i = 1; i <= 3; i++) {
                XSSFRow r = src.createRow(i);
                r.createCell(0).setCellValue("R" + i);
                r.createCell(1).setCellValue(i * 100);
            }
            XSSFSheet ps = wb.createSheet("Pivot");
            XSSFPivotTable pivot = ps.createPivotTable(
                    new AreaReference("A1:B4", SpreadsheetVersion.EXCEL2007),
                    new CellReference(0, 0), src);
            pivot.addRowLabel(0);
            pivot.addColumnLabel(DataConsolidateFunction.AVERAGE, 1);
            pivot.addColumnLabel(DataConsolidateFunction.MAX, 1);
            pivot.addColumnLabel(DataConsolidateFunction.MIN, 1);
            pivot.addColumnLabel(DataConsolidateFunction.COUNT, 1);
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return bos.toByteArray();
        }
    }

    private byte[] workbookWithoutPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Quarter");
            h.createCell(2).setCellValue("Amount");
            for (int i = 1; i <= 2; i++) {
                XSSFRow r = src.createRow(i);
                r.createCell(0).setCellValue(i == 1 ? "E" : "W");
                r.createCell(1).setCellValue("Q" + i);
                r.createCell(2).setCellValue(i * 100);
            }
            wb.createSheet("Pivot");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return bos.toByteArray();
        }
    }
}
