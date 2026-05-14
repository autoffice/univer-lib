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
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Pivot table converter tests: 覆盖读路径字段映射与写路径基础 best-effort 回写。
 */
class PivotTableConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_extract_row_column_value_fields_from_poi_pivot() throws Exception {
        byte[] bytes = buildWorkbookWithPivot();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = PivotTableConverter.readSheetPivotTables(wb, pivotSheet, mapper, "wb", "Pivot");

            assertThat(payload.size()).isEqualTo(1);
            JsonNode pivot = payload.get(0);
            assertThat(pivot.path("sourceRangeInfo").path("sheetName").asText()).isEqualTo("Source");
            assertThat(pivot.path("sourceRangeInfo").path("range").path("startRow").asInt()).isEqualTo(0);
            assertThat(pivot.path("sourceRangeInfo").path("range").path("endRow").asInt()).isEqualTo(2);
            assertThat(pivot.path("targetCellInfo").path("sheetName").asText()).isEqualTo("Pivot");
            JsonNode fieldsConfig = pivot.path("fieldsConfig");
            assertThat(fieldsConfig.path("rows").size()).isEqualTo(1);
            assertThat(fieldsConfig.path("rows").get(0).path("sourceName").asText()).isEqualTo("Region");
            assertThat(fieldsConfig.path("columns").size()).isEqualTo(1);
            assertThat(fieldsConfig.path("columns").get(0).path("sourceName").asText()).isEqualTo("Quarter");
            assertThat(fieldsConfig.path("values").size()).isEqualTo(1);
            assertThat(fieldsConfig.path("values").get(0).path("subtotal").asText()).isEqualTo("sum");

            JsonNode unsupported = pivot.path("unsupportedFeatures");
            assertThat(unsupported.isArray()).isTrue();
            java.util.Set<String> alwaysUnsupported = new java.util.LinkedHashSet<>();
            unsupported.forEach(n -> alwaysUnsupported.add(n.asText()));
            assertThat(alwaysUnsupported).contains(
                    "grouping",
                    "drillDown",
                    "refreshSemantics",
                    "advancedLabelSort",
                    "advancedLabelFilter",
                    "valuePositioning",
                    "executionMode");

            JsonNode detected = pivot.path("detectedFeatures");
            assertThat(detected.isArray()).isTrue();
        }
    }

    @Test
    void should_write_best_effort_pivot_back_from_payload() throws Exception {
        byte[] bytes = buildWorkbookWithoutPivot();

        ObjectNode pivotNode = mapper.createObjectNode();
        pivotNode.put("pivotTableId", "pivot-1");
        ObjectNode source = mapper.createObjectNode();
        source.put("sheetName", "Source");
        source.put("subUnitId", "Source");
        ObjectNode sourceRange = mapper.createObjectNode();
        sourceRange.put("startRow", 0);
        sourceRange.put("startColumn", 0);
        sourceRange.put("endRow", 2);
        sourceRange.put("endColumn", 2);
        source.set("range", sourceRange);
        pivotNode.set("sourceRangeInfo", source);

        ObjectNode target = mapper.createObjectNode();
        target.put("sheetName", "Pivot");
        target.put("row", 0);
        target.put("column", 0);
        pivotNode.set("targetCellInfo", target);

        ObjectNode fieldsConfig = mapper.createObjectNode();
        ArrayNode rows = mapper.createArrayNode();
        rows.add(mapper.createObjectNode().put("sourceIndex", 0).put("sourceName", "Region"));
        ArrayNode columns = mapper.createArrayNode();
        columns.add(mapper.createObjectNode().put("sourceIndex", 1).put("sourceName", "Quarter"));
        ArrayNode values = mapper.createArrayNode();
        values.add(mapper.createObjectNode().put("sourceIndex", 2).put("sourceName", "Amount")
                .put("subtotal", "sum"));
        fieldsConfig.set("rows", rows);
        fieldsConfig.set("columns", columns);
        fieldsConfig.set("filters", mapper.createArrayNode());
        fieldsConfig.set("values", values);
        pivotNode.set("fieldsConfig", fieldsConfig);
        ArrayNode payload = mapper.createArrayNode();
        payload.add(pivotNode);

        byte[] out;
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            Map<String, XSSFSheet> sheetMap = new LinkedHashMap<>();
            sheetMap.put("Source", wb.getSheet("Source"));
            sheetMap.put("Pivot", pivotSheet);
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, sheetMap);
            ByteArrayOutputStream buf = new ByteArrayOutputStream();
            wb.write(buf);
            out = buf.toByteArray();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out))) {
            assertThat(wb.getSheet("Pivot").getPivotTables()).hasSize(1);
        }
    }

    // ------------------------------------------------------------
    // helpers
    // ------------------------------------------------------------
    private byte[] buildWorkbookWithPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet source = wb.createSheet("Source");
            XSSFRow header = source.createRow(0);
            header.createCell(0).setCellValue("Region");
            header.createCell(1).setCellValue("Quarter");
            header.createCell(2).setCellValue("Amount");
            XSSFRow r1 = source.createRow(1);
            r1.createCell(0).setCellValue("East");
            r1.createCell(1).setCellValue("Q1");
            r1.createCell(2).setCellValue(100);
            XSSFRow r2 = source.createRow(2);
            r2.createCell(0).setCellValue("West");
            r2.createCell(1).setCellValue("Q1");
            r2.createCell(2).setCellValue(200);

            XSSFSheet pivotSheet = wb.createSheet("Pivot");
            XSSFPivotTable pivot = pivotSheet.createPivotTable(
                    new AreaReference("A1:C3", SpreadsheetVersion.EXCEL2007),
                    new CellReference(0, 0),
                    source);
            pivot.addRowLabel(0);
            pivot.addColLabel(1);
            pivot.addColumnLabel(DataConsolidateFunction.SUM, 2);

            ByteArrayOutputStream buf = new ByteArrayOutputStream();
            wb.write(buf);
            return buf.toByteArray();
        }
    }

    private byte[] buildWorkbookWithoutPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet source = wb.createSheet("Source");
            XSSFRow header = source.createRow(0);
            header.createCell(0).setCellValue("Region");
            header.createCell(1).setCellValue("Quarter");
            header.createCell(2).setCellValue("Amount");
            XSSFRow r1 = source.createRow(1);
            r1.createCell(0).setCellValue("East");
            r1.createCell(1).setCellValue("Q1");
            r1.createCell(2).setCellValue(100);
            XSSFRow r2 = source.createRow(2);
            r2.createCell(0).setCellValue("West");
            r2.createCell(1).setCellValue("Q2");
            r2.createCell(2).setCellValue(300);
            wb.createSheet("Pivot");
            ByteArrayOutputStream buf = new ByteArrayOutputStream();
            wb.write(buf);
            return buf.toByteArray();
        }
    }
}
