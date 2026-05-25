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
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 把所有 plugin resource 都放进 sidecar，跑一次 round-trip：
 * 让 WorkbookConverter 的 CF/DV/Notes/Pictures/Shapes/Charts/Sparkline/Filter/Tables/DefinedName
 * 各自的回写分支都执行一遍。
 */
class WorkbookConverterAllPluginsRoundTripTest {

    @Test
    void should_round_trip_every_plugin_resource() throws Exception {
        IWorkbookData wb = new IWorkbookData().setId("wb-all").setLocale("enUS");
        IWorksheetData s1 = new IWorksheetData().setId("s1").setName("Sheet1");
        IWorksheetData s2 = new IWorksheetData().setId("s2").setName("Sheet2");
        wb.getSheets().put("s1", s1);
        wb.getSheets().put("s2", s2);
        wb.setSheetOrder(Arrays.asList("s1", "s2"));

        List<Object> resources = new ArrayList<>();
        resources.add(plugin("SHEET_CONDITIONAL_FORMATTING_PLUGIN",
                "{\"s1\":[{\"cfId\":\"a\",\"ranges\":[{\"startRow\":0,\"startColumn\":0,\"endRow\":0,\"endColumn\":0}]," +
                        "\"rule\":{\"type\":\"highlightCell\",\"subType\":\"formula\",\"value\":\"=A1=1\"}}]}"));
        resources.add(plugin("SHEET_DATA_VALIDATION_PLUGIN",
                "{\"s1\":{\"r1\":{\"uid\":\"r1\",\"type\":\"list\"," +
                        "\"ranges\":[{\"startRow\":0,\"startColumn\":0,\"endRow\":0,\"endColumn\":0}]," +
                        "\"formula1\":\"a,b,c\"}}}"));
        resources.add(plugin("SHEET_NOTE_PLUGIN",
                "{\"s1\":{\"0\":{\"0\":{\"note\":\"hello\",\"width\":160,\"height\":100,\"show\":false}}}}"));
        // 图片：1×1 PNG
        String png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
        resources.add(plugin("SHEET_DRAWING_PLUGIN",
                "{\"s1\":{\"data\":{\"d1\":{\"drawingType\":0,\"imageSourceType\":\"BASE64\",\"source\":\"" + png +
                        "\",\"sheetTransform\":{\"from\":{\"column\":0,\"columnOffset\":0,\"row\":0,\"rowOffset\":0}," +
                        "\"to\":{\"column\":2,\"columnOffset\":0,\"row\":2,\"rowOffset\":0}}}},\"order\":[\"d1\"]}}"));
        resources.add(plugin("SHEET_FILTER_PLUGIN",
                "{\"s1\":{\"ref\":\"A1:B2\"}}"));
        resources.add(plugin("SHEET_TABLE_PLUGIN",
                "{\"s1\":{\"t1\":{\"name\":\"Tbl\",\"ref\":\"A1:B3\"}}}"));
        // workbook-level defined name
        resources.add(plugin("SHEET_DEFINED_NAME_PLUGIN",
                "{\"id1\":{\"id\":\"id1\",\"name\":\"MyName\",\"formulaOrRefString\":\"Sheet1!$A$1\"}}"));
        // chart 与 sparkline 走 rawXml 占位（即使非法也不应阻塞写入）
        resources.add(plugin("SHEET_CHART_PLUGIN",
                "{\"s1\":{\"data\":{\"c1\":{\"rawXml\":\"<not-valid/>\"}},\"order\":[\"c1\"]}}"));
        resources.add(plugin("SHEET_SHAPE_PLUGIN",
                "{\"s1\":{\"data\":{\"sh1\":{\"shapeType\":\"rect\",\"rawXml\":\"<not-valid/>\"}}," +
                        "\"order\":[\"sh1\"]}}"));
        resources.add(plugin("SHEET_SPARKLINE_PLUGIN",
                "{\"s1\":{\"rawXml\":\"<not-valid/>\"}}"));
        // 透视表：缺信息 → 应被跳过
        resources.add(plugin("SHEET_PIVOT_TABLE_PLUGIN",
                "{\"s1\":[{\"pivotTableId\":\"p1\"}]}"));
        wb.setResources(resources);

        // 让 Sheet1 至少有几个单元格，便于 table / 图片落点
        Map<Integer, Map<Integer, io.github.autoffice.univer.model.ICellData>> cells = new LinkedHashMap<>();
        Map<Integer, io.github.autoffice.univer.model.ICellData> r0 = new LinkedHashMap<>();
        r0.put(0, new io.github.autoffice.univer.model.ICellData().setV("Header1"));
        r0.put(1, new io.github.autoffice.univer.model.ICellData().setV("Header2"));
        cells.put(0, r0);
        Map<Integer, io.github.autoffice.univer.model.ICellData> r1 = new LinkedHashMap<>();
        r1.put(0, new io.github.autoffice.univer.model.ICellData().setV("a"));
        r1.put(1, new io.github.autoffice.univer.model.ICellData().setV("b"));
        cells.put(1, r1);
        Map<Integer, io.github.autoffice.univer.model.ICellData> r2 = new LinkedHashMap<>();
        r2.put(0, new io.github.autoffice.univer.model.ICellData().setV("c"));
        r2.put(1, new io.github.autoffice.univer.model.ICellData().setV("d"));
        cells.put(2, r2);
        s1.setCellData(cells);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out, UniverXlsxOptions.builder().writeSidecar(true).prettyJson(true).build());
        // 不抛错
        assertThat(out.size()).isPositive();

        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back).isNotNull();
        assertThat(back.getSheets()).containsKey("s1");
        // resources 仍存在
        assertThat(back.getResources()).isInstanceOf(List.class);
    }

    private Map<String, Object> plugin(String name, String dataJson) {
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", name);
        entry.put("data", dataJson);
        return entry;
    }
}
