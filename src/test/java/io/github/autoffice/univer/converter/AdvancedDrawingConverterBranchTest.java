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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * AdvancedDrawingConverter 边角分支：所有 shapeType / 空 payload / 坏 rawXml / 跳过其它图形。
 */
class AdvancedDrawingConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_handle_every_shape_type_round_trip() throws Exception {
        // shapeTypeCodeOf + shapeTypeNameOf 全分支
        String[] kinds = {"shape", "shape", "shape", "shape", "shape", "shape", "shape", "shape",
                "connector", "connector", "connector", "connector", "connector", "shape"};
        String[] types = {"rect", "roundRect", "diamond", "smileyFace",
                "flowChartProcess", "flowChartDecision", "flowChartTerminator", "flowChartDocument",
                "straightConnector1", "bentConnector2", "bentConnector3",
                "curvedConnector2", "curvedConnector3", "downArrow"};
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        ArrayNode order = mapper.createArrayNode();
        for (int i = 0; i < types.length; i++) {
            data.set("s" + i, shapeNode("s" + i, kinds[i], types[i],
                    i, 0, i, 0, i + 2, 0, i + 2, 0));
            order.add("s" + i);
        }
        payload.set("data", data);
        payload.set("order", order);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            AdvancedDrawingConverter.writeSheetShapes(sh, payload);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                ObjectNode reread = AdvancedDrawingConverter.readSheetShapes(
                        wb2.getSheetAt(0), mapper, "u", "S");
                // 应能读出全部
                assertThat(reread.path("data").size()).isEqualTo(types.length);
            }
        }
    }

    @Test
    void should_skip_chart_with_empty_or_invalid_rawXml() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            // 空 rawXml → 应跳过，不创建 chart
            data.set("c1", chartItem(""));
            // 非法 rawXml → 创建空 chart 容器但不抛错
            data.set("c2", chartItem("<not valid"));
            payload.set("data", data);
            ArrayNode order = mapper.createArrayNode();
            order.add("c1");
            order.add("c2");
            payload.set("order", order);
            AdvancedDrawingConverter.writeSheetCharts(sh, payload);
            // c1 跳过，c2 创建空容器
            assertThat(sh.getDrawingPatriarch().getCharts()).hasSize(1);
        }
    }

    @Test
    void should_skip_chart_when_payload_invalid() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            AdvancedDrawingConverter.writeSheetCharts(sh, null);
            AdvancedDrawingConverter.writeSheetCharts(sh, mapper.createArrayNode());
            ObjectNode noData = mapper.createObjectNode();
            AdvancedDrawingConverter.writeSheetCharts(sh, noData);
            ObjectNode arrayData = mapper.createObjectNode();
            arrayData.set("data", mapper.createArrayNode());
            AdvancedDrawingConverter.writeSheetCharts(sh, arrayData);
            // 空 data
            ObjectNode emptyData = mapper.createObjectNode();
            emptyData.set("data", mapper.createObjectNode());
            AdvancedDrawingConverter.writeSheetCharts(sh, emptyData);
            // 验证没有意外 NPE
        }
    }

    @Test
    void should_return_empty_chart_payload_for_null_or_no_drawing() throws Exception {
        ObjectNode out1 = AdvancedDrawingConverter.readSheetCharts(null, mapper, "u", "s");
        assertThat(out1.path("data").size()).isEqualTo(0);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode out2 = AdvancedDrawingConverter.readSheetCharts(sh, mapper, "u", "S");
            assertThat(out2.path("data").size()).isEqualTo(0);
        }
    }

    @Test
    void should_return_empty_shape_payload_for_null_or_no_drawing() throws Exception {
        ObjectNode out1 = AdvancedDrawingConverter.readSheetShapes(null, mapper, "u", "s");
        assertThat(out1.path("data").size()).isEqualTo(0);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode out2 = AdvancedDrawingConverter.readSheetShapes(sh, mapper, "u", "S");
            assertThat(out2.path("data").size()).isEqualTo(0);
        }
    }

    @Test
    void should_return_empty_sparkline_for_null_sheet_or_no_extLst() throws Exception {
        ObjectNode out1 = AdvancedDrawingConverter.readSheetSparkline(null, mapper);
        assertThat(out1).isEmpty();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode out2 = AdvancedDrawingConverter.readSheetSparkline(sh, mapper);
            assertThat(out2).isEmpty();
        }
    }

    @Test
    void should_skip_sparkline_with_empty_or_invalid_xml() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // null payload
            AdvancedDrawingConverter.writeSheetSparkline(sh, null);
            // 非 object
            AdvancedDrawingConverter.writeSheetSparkline(sh, mapper.createArrayNode());
            // 空 extLstXml
            AdvancedDrawingConverter.writeSheetSparkline(sh, mapper.createObjectNode());
            // 非法 xml
            ObjectNode badXml = mapper.createObjectNode();
            badXml.put("extLstXml", "<not valid");
            AdvancedDrawingConverter.writeSheetSparkline(sh, badXml);
            // 不抛错
            assertThat(sh.getCTWorksheet().isSetExtLst()).isFalse();
        }
    }

    @Test
    void should_skip_invalid_shape_payload() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            AdvancedDrawingConverter.writeSheetShapes(sh, null);
            AdvancedDrawingConverter.writeSheetShapes(sh, mapper.createArrayNode());
            ObjectNode noData = mapper.createObjectNode();
            AdvancedDrawingConverter.writeSheetShapes(sh, noData);
            ObjectNode arrayData = mapper.createObjectNode();
            arrayData.set("data", mapper.createArrayNode());
            AdvancedDrawingConverter.writeSheetShapes(sh, arrayData);
        }
    }

    private ObjectNode chartItem(String rawXml) {
        ObjectNode item = mapper.createObjectNode();
        item.put("rawXml", rawXml);
        ObjectNode st = mapper.createObjectNode();
        ObjectNode from = mapper.createObjectNode();
        from.put("column", 0);
        from.put("columnOffset", 0);
        from.put("row", 0);
        from.put("rowOffset", 0);
        ObjectNode to = mapper.createObjectNode();
        to.put("column", 5);
        to.put("columnOffset", 0);
        to.put("row", 5);
        to.put("rowOffset", 0);
        st.set("from", from);
        st.set("to", to);
        item.set("sheetTransform", st);
        return item;
    }

    private ObjectNode shapeNode(String id, String kind, String shapeType,
                                 int fromCol, int fromColOff, int fromRow, int fromRowOff,
                                 int toCol, int toColOff, int toRow, int toRowOff) {
        ObjectNode n = mapper.createObjectNode();
        n.put("unitId", "u");
        n.put("subUnitId", "S");
        n.put("shapeId", id);
        n.put("kind", kind);
        n.put("shapeType", shapeType);
        ObjectNode from = mapper.createObjectNode();
        from.put("column", fromCol);
        from.put("columnOffset", fromColOff);
        from.put("row", fromRow);
        from.put("rowOffset", fromRowOff);
        ObjectNode to = mapper.createObjectNode();
        to.put("column", toCol);
        to.put("columnOffset", toColOff);
        to.put("row", toRow);
        to.put("rowOffset", toRowOff);
        ObjectNode st = mapper.createObjectNode();
        st.set("from", from);
        st.set("to", to);
        n.set("sheetTransform", st);
        return n;
    }
}
