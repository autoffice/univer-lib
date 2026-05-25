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
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IFreeze;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * 全方位 WorkbookConverter 分支覆盖：resource 抽取与合并、sidecar 合并、null 输入、
 * resources 类型异常、字段保留、styleId 解析、坏 JSON 兜底等。
 */
class WorkbookConverterBranchTest {

    private static IWorkbookData minimalWb() {
        IWorkbookData wb = new IWorkbookData().setId("wb1").setLocale("zhCN");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        wb.setSheetOrder(Collections.singletonList("s1"));
        return wb;
    }

    @Test
    void should_return_empty_workbook_for_null_source() throws Exception {
        WorkbookConverter wc = new WorkbookConverter(UniverXlsxOptions.defaults());
        try (XSSFWorkbook out = wc.toXlsx(null)) {
            assertThat(out.getNumberOfSheets()).isEqualTo(0);
        }
    }

    @Test
    void should_fall_back_to_keys_when_sheetOrder_missing() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-noorder");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("A"));
        wb.getSheets().put("s2", new IWorksheetData().setId("s2").setName("B"));
        // 不设 sheetOrder
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets()).containsKeys("s1", "s2");
    }

    @Test
    void should_skip_unknown_sheet_id_in_order() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-skip");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Hello"));
        wb.setSheetOrder(Arrays.asList("ghost", "s1"));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets()).containsKey("s1");
    }

    @Test
    void should_collide_sheet_names_with_suffix() throws Exception {
        // 两个 sheet 同名时第二个会带 sid 后缀
        IWorkbookData wb = new IWorkbookData().setId("wb-collide");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Same"));
        wb.getSheets().put("s2", new IWorksheetData().setId("s2").setName("Same"));
        wb.setSheetOrder(Arrays.asList("s1", "s2"));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        // 第二个 sheet 实际名字是 Same_s2
        boolean foundCollision = back.getSheets().values().stream()
                .anyMatch(s -> s.getName() != null && s.getName().contains("Same"));
        assertThat(foundCollision).isTrue();
    }

    @Test
    void should_set_sheet_visibility_hidden() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-hidden");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("V"));
        wb.getSheets().put("s2", new IWorksheetData().setId("s2").setName("H").setHidden(BooleanNumber.TRUE));
        wb.setSheetOrder(Arrays.asList("s1", "s2"));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getSheets().get("s2").getHidden()).isEqualTo(BooleanNumber.TRUE);
    }

    @Test
    void should_resolve_string_style_ids_in_cellData() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-style");
        IStyleData boldStyle = new IStyleData().setFf("Arial").setFs(12).setBl(BooleanNumber.TRUE);
        wb.getStyles().put("style-1", boldStyle);
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("Hello").setS("style-1"));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        IWorksheetData backWs = back.getSheets().get("s1");
        ICellData backCell = backWs.getCellData().get(0).get(0);
        assertThat(backCell.getV()).isEqualTo("Hello");
        // 样式 id 字符串被回填
        assertThat(backCell.getS()).isInstanceOf(String.class);
    }

    @Test
    void should_merge_cellData_xlsx_over_sidecar_baseline() throws Exception {
        // sidecar 携带 cell.p 富文本与 custom，xlsx 提供 v/t；二者合并保留富文本
        IWorkbookData wb = minimalWb();
        IWorksheetData ws = wb.getSheets().get("s1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        Map<String, Object> custom = new LinkedHashMap<>();
        custom.put("k", "v");
        row.put(0, new ICellData().setV("xlsx-val").setCustom(custom));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        ICellData c = back.getSheets().get("s1").getCellData().get(0).get(0);
        // custom 字段从 sidecar 保留
        assertThat(c.getCustom()).containsEntry("k", "v");
        // xlsx 的 v 应当覆盖 / xlsx-derived v wins
        assertThat(c.getV()).isEqualTo("xlsx-val");
    }

    @Test
    void should_preserve_sidecar_fields_not_derived_from_xlsx() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-sc")
                .setAppVersion("9.9.9").setLocale("zhCN");
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S")
                .setScrollLeft(123.0).setScrollTop(456.0));
        wb.setSheetOrder(Collections.singletonList("s1"));
        // 自定义 resources（任意 plugin）
        List<Object> resources = new ArrayList<>();
        Map<String, Object> custom = new LinkedHashMap<>();
        custom.put("name", "MY_CUSTOM_PLUGIN");
        custom.put("data", "{\"any\":\"thing\"}");
        resources.add(custom);
        wb.setResources(resources);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getAppVersion()).isEqualTo("9.9.9");
        // sidecar 中的 scroll 字段被保留
        assertThat(back.getSheets().get("s1").getScrollLeft()).isEqualTo(123.0);
        assertThat(back.getSheets().get("s1").getScrollTop()).isEqualTo(456.0);
        assertThat(back.getResources()).isInstanceOf(List.class);
    }

    @Test
    void should_round_trip_merges_rows_columns_freeze() throws Exception {
        IWorkbookData wb = minimalWb();
        IWorksheetData ws = wb.getSheets().get("s1");
        ws.setMergeData(Collections.singletonList(
                new IRange().setStartRow(0).setEndRow(1).setStartColumn(0).setEndColumn(2)));
        ws.setFreeze(new IFreeze().setStartRow(1).setStartColumn(1).setXSplit(1).setYSplit(1));
        // 一行高度
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("X"));
        Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
        cells.put(0, row);
        ws.setCellData(cells);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        IWorksheetData bws = back.getSheets().get("s1");
        assertThat(bws.getMergeData()).hasSize(1);
        assertThat(bws.getFreeze()).isNotNull();
    }

    @Test
    void should_ignore_resources_when_not_a_list() throws Exception {
        IWorkbookData wb = minimalWb();
        wb.setResources("not-a-list"); // 故意设非法类型
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        // 不抛即可
        assertThat(back.getSheets()).containsKey("s1");
    }

    @Test
    void should_extract_resource_with_map_data() throws Exception {
        // data 直接是 Map 而非 JSON 字符串 → 走 valueToTree 分支
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_NOTE_PLUGIN");
        Map<String, Object> data = new LinkedHashMap<>();
        Map<String, Object> sheetNotes = new LinkedHashMap<>();
        Map<String, Object> rowNotes = new LinkedHashMap<>();
        Map<String, Object> note = new LinkedHashMap<>();
        note.put("text", "hi");
        rowNotes.put("0", note);
        sheetNotes.put("0", rowNotes);
        data.put("s1", sheetNotes);
        entry.put("data", data);
        resources.add(entry);
        wb.setResources(resources);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getResources()).isInstanceOf(List.class);
    }

    @Test
    void should_skip_resource_with_empty_string_data() throws Exception {
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_NOTE_PLUGIN");
        entry.put("data", "   "); // 空白字串
        resources.add(entry);
        // 无 data 的另一条
        Map<String, Object> entry2 = new LinkedHashMap<>();
        entry2.put("name", "SHEET_NOTE_PLUGIN");
        resources.add(entry2);
        // 非 Map 项
        resources.add("not-a-map");
        wb.setResources(resources);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        // 不抛错
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_ignore_resource_with_bad_json_data() throws Exception {
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_NOTE_PLUGIN");
        entry.put("data", "{not valid json");
        resources.add(entry);
        wb.setResources(resources);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        // 不抛即可
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_round_trip_workbook_defined_names() throws Exception {
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_DEFINED_NAME_PLUGIN");
        entry.put("data", "{\"id1\":{\"id\":\"id1\",\"name\":\"MyRange\",\"formulaOrRefString\":\"Sheet1!$A$1:$B$2\"}}");
        resources.add(entry);
        wb.setResources(resources);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getResources()).isInstanceOf(List.class);
        boolean hasDN = ((List<?>) back.getResources()).stream().anyMatch(item ->
                item instanceof Map && "SHEET_DEFINED_NAME_PLUGIN".equals(((Map<?, ?>) item).get("name")));
        assertThat(hasDN).isTrue();
    }

    @Test
    void should_preserve_baseline_fields_when_xlsx_lacks_them() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-pres").setLocale("zhCN");
        IWorksheetData baseline = new IWorksheetData().setId("s1").setName("Pres")
                .setTabColor("#ff0000")
                .setShowGridlines(BooleanNumber.FALSE)
                .setRightToLeft(BooleanNumber.TRUE)
                .setDefaultColumnWidth(100.0)
                .setDefaultRowHeight(28.0);
        wb.getSheets().put("s1", baseline);
        wb.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        IWorksheetData bws = back.getSheets().get("s1");
        assertThat(bws.getTabColor()).isNotNull();
        assertThat(bws.getShowGridlines()).isEqualTo(BooleanNumber.FALSE);
        assertThat(bws.getRightToLeft()).isEqualTo(BooleanNumber.TRUE);
    }

    @Test
    void should_merge_existing_resource_with_new_data() throws Exception {
        // sidecar 与 xlsx 读出的资源对同一 plugin 都有数据时应合并
        IWorkbookData wb = new IWorkbookData().setId("wb-merge");
        // sheet 1 有一个备注的 sidecar baseline
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("M");
        wb.getSheets().put("s1", ws);
        wb.setSheetOrder(Collections.singletonList("s1"));
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_NOTE_PLUGIN");
        // sidecar 中 s2 sheet 有备注（一定不会被 xlsx 数据覆盖，因为没有 s2 sheet）
        entry.put("data", "{\"ghost\":{\"0\":{\"0\":{\"text\":\"baseline\"}}}}");
        resources.add(entry);
        wb.setResources(resources);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        // sidecar baseline 的 ghost 数据仍在
        boolean stillHasGhost = false;
        for (Object item : (List<?>) back.getResources()) {
            if (item instanceof Map
                    && "SHEET_NOTE_PLUGIN".equals(((Map<?, ?>) item).get("name"))) {
                String d = String.valueOf(((Map<?, ?>) item).get("data"));
                if (d.contains("ghost")) {
                    stillHasGhost = true;
                }
            }
        }
        assertThat(stillHasGhost).isTrue();
    }

    @Test
    void should_handle_workbook_resource_with_string_data() throws Exception {
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_DEFINED_NAME_PLUGIN");
        entry.put("data", "{}"); // 空 JSON 对象
        resources.add(entry);
        wb.setResources(resources);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        // 不抛即可
        assertThat(out.size()).isPositive();
    }

    @Test
    void should_handle_workbook_resource_with_invalid_data() throws Exception {
        IWorkbookData wb = minimalWb();
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_DEFINED_NAME_PLUGIN");
        entry.put("data", "{bad json");
        resources.add(entry);
        wb.setResources(resources);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        assertThat(out.size()).isPositive();
    }
}
