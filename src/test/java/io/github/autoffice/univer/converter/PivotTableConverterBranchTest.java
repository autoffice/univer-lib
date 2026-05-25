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

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
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
 * Pivot 边角分支：null/invalid payload、不同 subtotal、options、filters、displayName。
 */
class PivotTableConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_skip_writing_when_inputs_are_null() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // null sheetPayload
            PivotTableConverter.writeSheetPivotTables(wb, sh, null, Collections.emptyMap());
            // null wb
            PivotTableConverter.writeSheetPivotTables(null, sh, mapper.createArrayNode(), Collections.emptyMap());
            // null sheet
            PivotTableConverter.writeSheetPivotTables(wb, null, mapper.createArrayNode(), Collections.emptyMap());
            // 非 array
            PivotTableConverter.writeSheetPivotTables(wb, sh, mapper.createObjectNode(), Collections.emptyMap());
            assertThat(sh.getPivotTables()).isEmpty();
        }
    }

    @Test
    void should_return_empty_for_null_sheet_on_read() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            ArrayNode out = PivotTableConverter.readSheetPivotTables(wb, null, mapper, "u", "s");
            assertThat(out).isEmpty();
        }
    }

    @Test
    void should_skip_pivot_with_missing_range_or_target() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            // 缺 sourceRangeInfo.range
            ArrayNode bad = mapper.createArrayNode();
            ObjectNode b1 = mapper.createObjectNode();
            b1.put("pivotTableId", "p1");
            b1.set("sourceRangeInfo", mapper.createObjectNode());
            b1.set("targetCellInfo", mapper.createObjectNode());
            b1.set("fieldsConfig", mapper.createObjectNode());
            bad.add(b1);
            // 缺 targetCellInfo
            ObjectNode b2 = mapper.createObjectNode();
            b2.set("sourceRangeInfo",
                    mapper.createObjectNode().set("range", rangeNode()));
            b2.put("missingTarget", true);
            bad.add(b2);
            // 缺 fieldsConfig
            ObjectNode b3 = mapper.createObjectNode();
            b3.set("sourceRangeInfo",
                    mapper.createObjectNode().set("range", rangeNode()));
            b3.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
            bad.add(b3);
            // sourceSheet 不存在
            ObjectNode b4 = mapper.createObjectNode();
            b4.set("sourceRangeInfo",
                    mapper.createObjectNode().set("range", rangeNode()));
            b4.put("sourceRangeInfo.sheetName", "ghost");
            b4.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
            b4.set("fieldsConfig", mapper.createObjectNode());
            bad.add(b4);

            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, bad, Collections.emptyMap());
            assertThat(pivotSheet.getPivotTables()).isEmpty();
        }
    }

    @Test
    void should_skip_pivot_with_invalid_source_range() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode bad = mapper.createArrayNode();
            ObjectNode b = mapper.createObjectNode();
            ObjectNode src = mapper.createObjectNode();
            src.put("sheetName", "Source");
            // 缺范围字段
            ObjectNode badRange = mapper.createObjectNode();
            badRange.put("startRow", -1);
            src.set("range", badRange);
            b.set("sourceRangeInfo", src);
            b.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
            b.set("fieldsConfig", mapper.createObjectNode());
            bad.add(b);

            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, bad, map);
            assertThat(pivotSheet.getPivotTables()).isEmpty();
        }
    }

    @Test
    void should_skip_pivot_with_invalid_target() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode bad = mapper.createArrayNode();
            ObjectNode b = mapper.createObjectNode();
            ObjectNode src = mapper.createObjectNode();
            src.put("sheetName", "Source");
            src.set("range", rangeNode());
            b.set("sourceRangeInfo", src);
            b.set("targetCellInfo",
                    mapper.createObjectNode().put("row", -1).put("column", 0));
            b.set("fieldsConfig", mapper.createObjectNode());
            bad.add(b);

            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, bad, map);
            assertThat(pivotSheet.getPivotTables()).isEmpty();
        }
    }

    @Test
    void should_apply_options_and_name() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = mapper.createObjectNode();
            p.put("pivotTableId", "px");
            p.put("name", "MyPivot");
            ObjectNode src = mapper.createObjectNode();
            src.put("sheetName", "Source");
            src.set("range", rangeNode());
            p.set("sourceRangeInfo", src);
            p.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
            ObjectNode fc = mapper.createObjectNode();
            ArrayNode rows = mapper.createArrayNode();
            rows.add(mapper.createObjectNode().put("sourceIndex", 0).put("sourceName", "Region"));
            ArrayNode columns = mapper.createArrayNode();
            // 带 displayName
            columns.add(mapper.createObjectNode().put("sourceIndex", 1).put("sourceName", "Quarter")
                    .put("displayName", "Q"));
            ArrayNode filters = mapper.createArrayNode();
            filters.add(mapper.createObjectNode().put("sourceIndex", 0));
            ArrayNode values = mapper.createArrayNode();
            values.add(mapper.createObjectNode().put("sourceIndex", 2).put("subtotal", "average")
                    .put("displayName", "Avg"));
            fc.set("rows", rows);
            fc.set("columns", columns);
            fc.set("filters", filters);
            fc.set("values", values);
            p.set("fieldsConfig", fc);

            // 各种 options
            ObjectNode opts = mapper.createObjectNode();
            opts.put("showRowGrandTotal", false);
            opts.put("showColumnGrandTotal", false);
            opts.put("compact", false);
            opts.put("outline", true);
            opts.put("outlineData", true);
            opts.put("compactData", false);
            opts.put("multipleFieldFilters", true);
            opts.put("dataCaption", "Vals");
            opts.put("dataOnRows", true);
            p.set("options", opts);

            payload.add(p);

            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("Source", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
            // name 应被设置
            assertThat(pivotSheet.getPivotTables().get(0).getCTPivotTableDefinition().getName())
                    .isEqualTo("MyPivot");
        }
    }

    @Test
    void should_resolve_source_via_subUnitId_lookup() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            ObjectNode p = mapper.createObjectNode();
            ObjectNode src = mapper.createObjectNode();
            // 空 sheetName，但 subUnitId 命中
            src.put("subUnitId", "sourceSid");
            src.set("range", rangeNode());
            p.set("sourceRangeInfo", src);
            p.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
            p.set("fieldsConfig", mapper.createObjectNode());
            payload.add(p);

            Map<String, XSSFSheet> map = new LinkedHashMap<>();
            map.put("sourceSid", wb.getSheet("Source"));
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
            assertThat(pivotSheet.getPivotTables()).hasSize(1);
        }
    }

    @Test
    void should_handle_all_subtotal_functions() throws Exception {
        String[] subs = {"count", "countNumbers", "average", "max", "min",
                         "product", "stdDev", "stdDevp", "var", "varp", "unknown"};
        for (String sub : subs) {
            byte[] bytes = workbookWithoutPivot();
            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                XSSFSheet pivotSheet = wb.getSheet("Pivot");
                ArrayNode payload = mapper.createArrayNode();
                ObjectNode p = mapper.createObjectNode();
                ObjectNode src = mapper.createObjectNode();
                src.put("sheetName", "Source");
                src.set("range", rangeNode());
                p.set("sourceRangeInfo", src);
                p.set("targetCellInfo", mapper.createObjectNode().put("row", 0).put("column", 0));
                ObjectNode fc = mapper.createObjectNode();
                ArrayNode values = mapper.createArrayNode();
                values.add(mapper.createObjectNode().put("sourceIndex", 2).put("subtotal", sub));
                fc.set("values", values);
                p.set("fieldsConfig", fc);
                payload.add(p);

                Map<String, XSSFSheet> map = new LinkedHashMap<>();
                map.put("Source", wb.getSheet("Source"));
                PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, map);
                assertThat(pivotSheet.getPivotTables()).as("subtotal=%s", sub).hasSize(1);
            }
        }
    }

    @Test
    void should_skip_invalid_pivot_node() throws Exception {
        byte[] bytes = workbookWithoutPivot();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet pivotSheet = wb.getSheet("Pivot");
            ArrayNode payload = mapper.createArrayNode();
            payload.add(mapper.nullNode());
            payload.add(mapper.createArrayNode()); // 非 object
            PivotTableConverter.writeSheetPivotTables(wb, pivotSheet, payload, Collections.emptyMap());
            assertThat(pivotSheet.getPivotTables()).isEmpty();
        }
    }

    private ObjectNode rangeNode() {
        ObjectNode range = mapper.createObjectNode();
        range.put("startRow", 0);
        range.put("startColumn", 0);
        range.put("endRow", 2);
        range.put("endColumn", 2);
        return range;
    }

    private byte[] workbookWithoutPivot() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet src = wb.createSheet("Source");
            XSSFRow h = src.createRow(0);
            h.createCell(0).setCellValue("Region");
            h.createCell(1).setCellValue("Quarter");
            h.createCell(2).setCellValue("Amount");
            XSSFRow r1 = src.createRow(1);
            r1.createCell(0).setCellValue("E");
            r1.createCell(1).setCellValue("Q1");
            r1.createCell(2).setCellValue(10);
            XSSFRow r2 = src.createRow(2);
            r2.createCell(0).setCellValue("W");
            r2.createCell(1).setCellValue("Q2");
            r2.createCell(2).setCellValue(20);
            wb.createSheet("Pivot");
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            return out.toByteArray();
        }
    }
}
