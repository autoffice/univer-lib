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

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * WorkbookConverter partial 分支精确补全：
 * - sheet 名为 null fallback
 * - resource 回写到不存在的 sheetId（sh == null 分支）
 * - styleId 不在 styles map 中 / cellData S 不是字符串
 * - 仅含某一种 fromXlsx 字段（mergeData/rowData/columnData）
 */
class WorkbookConverterPartialBranchTest {

    @Test
    void should_use_sid_as_sheetName_when_name_is_null() throws Exception {
        // 覆盖 L125: ws.getName() == null → 使用 sid 作为 sheetName
        IWorkbookData wb = new IWorkbookData().setId("w");
        IWorksheetData ws = new IWorksheetData().setId("s1"); // 不设 name
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        // sheet 名应该是 sid (s1)
        assertThat(back.getSheets()).containsKey("s1");
    }

    @Test
    void should_skip_resource_writeback_when_sheetId_not_in_workbook() throws Exception {
        // 覆盖 L156/165/174/192/201/210/219/228/237: sh == null 分支
        // 即 sidecar 的 resource 包含某个 sheetId，但 workbook 里没有对应的 sheet
        IWorkbookData wb = new IWorkbookData().setId("w");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S1"));
        wb.setSheetOrder(Collections.singletonList("s1"));
        // resources 中包含 ghost sheetId
        List<Object> resources = new ArrayList<>();
        // CF 资源指向不存在的 ghost sheetId
        Map<String, Object> cfRes = new LinkedHashMap<>();
        cfRes.put("name", "SHEET_CONDITIONAL_FORMATTING_PLUGIN");
        cfRes.put("data", "{\"ghost\":[{\"cfId\":\"a\",\"ranges\":[]," +
                "\"rule\":{\"type\":\"highlightCell\",\"subType\":\"formula\",\"value\":\"=A1=1\"}}]}");
        resources.add(cfRes);
        // Note 资源指向 ghost
        Map<String, Object> noteRes = new LinkedHashMap<>();
        noteRes.put("name", "SHEET_NOTE_PLUGIN");
        noteRes.put("data", "{\"ghost\":{\"0\":{\"0\":{\"note\":\"hi\"}}}}");
        resources.add(noteRes);
        // Filter 资源指向 ghost
        Map<String, Object> filterRes = new LinkedHashMap<>();
        filterRes.put("name", "SHEET_FILTER_PLUGIN");
        filterRes.put("data", "{\"ghost\":{\"ref\":\"A1:B2\"}}");
        resources.add(filterRes);
        wb.setResources(resources);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        // 不应抛错
        UniverXlsx.write(wb, out);
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_handle_only_mergeData_present_in_xlsx() throws Exception {
        // 覆盖 mergeSheetData 中只有 mergeData 的分支
        IWorkbookData wb = new IWorkbookData().setId("w");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        ws.setMergeData(Collections.singletonList(
                new io.github.autoffice.univer.model.IRange()
                        .setStartRow(0).setEndRow(1)
                        .setStartColumn(0).setEndColumn(1)));
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s1").getMergeData()).hasSize(1);
    }

    @Test
    void should_handle_styleId_not_found_in_styles_map() throws Exception {
        // 覆盖 L455: styleId 引用了不存在于 styles map 中的 id
        IWorkbookData wb = new IWorkbookData().setId("w");
        // 不在 styles 中注册 "missing-id"
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("x").setS("missing-id"));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        // 不抛错；cell value 仍写入
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s1").getCellData().get(0).get(0).getV()).isEqualTo("x");
    }

    @Test
    void should_handle_cellData_with_non_string_s_field() throws Exception {
        // 覆盖 L453: cd.getS() 不是 String（例如已经是 IStyleData 内联）
        IWorkbookData wb = new IWorkbookData().setId("w");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        // 内联 IStyleData 而非字符串 id
        row.put(0, new ICellData().setV("x").setS(new IStyleData().setFf("Arial")));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s1").getCellData().get(0).get(0).getV()).isEqualTo("x");
    }

    @Test
    void should_handle_null_cellData_in_resolveStyles() throws Exception {
        // 覆盖 L447: ws.getCellData() == null 分支（直接返回 ws，不做样式解引用）
        IWorkbookData wb = new IWorkbookData().setId("w");
        IStyleData boldStyle = new IStyleData().setFf("Arial").setFs(12);
        wb.getStyles().put("style-1", boldStyle);

        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        // 不设 cellData → 触发 ws.getCellData() == null 分支
        ws.setCellData(null);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_handle_only_rowData_in_fromXlsx_merge() throws Exception {
        // 覆盖 L472: rowData 非空分支
        IWorkbookData wb = new IWorkbookData().setId("w");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, io.github.autoffice.univer.model.IRowData> rowData = new LinkedHashMap<>();
        rowData.put(0, new io.github.autoffice.univer.model.IRowData().setH(25.0));
        ws.setRowData(rowData);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s1").getRowData()).isNotEmpty();
    }

    @Test
    void should_handle_only_columnData_in_fromXlsx_merge() throws Exception {
        // 覆盖 L475: columnData 非空分支
        IWorkbookData wb = new IWorkbookData().setId("w");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, io.github.autoffice.univer.model.IColumnData> colData = new LinkedHashMap<>();
        colData.put(0, new io.github.autoffice.univer.model.IColumnData().setW(120.0));
        ws.setColumnData(colData);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s1").getColumnData()).isNotEmpty();
    }

    @Test
    void should_handle_empty_styles_in_resolveStyles() throws Exception {
        // 覆盖 L444: wb.getStyles() == null || isEmpty 分支
        IWorkbookData wb = new IWorkbookData().setId("w");
        wb.setStyles(null); // null styles map

        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("x").setS("some-id"));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_use_existing_sheetOrder_when_present_in_baseline() throws Exception {
        // 覆盖 L425: out.getSheetOrder() != null 的分支
        IWorkbookData wb = new IWorkbookData().setId("w");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S1"));
        wb.getSheets().put("s2", new IWorksheetData().setId("s2").setName("S2"));
        wb.setSheetOrder(Arrays.asList("s1", "s2"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        // sheetOrder 应保留
        assertThat(back.getSheetOrder()).containsExactly("s1", "s2");
    }
}
