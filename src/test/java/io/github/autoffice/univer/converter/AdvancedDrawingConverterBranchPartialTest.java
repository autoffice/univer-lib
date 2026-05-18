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
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 针对 AdvancedDrawingConverter 中 partial 分支（null 检查/非 object/order 缺失等）的精确补全。
 */
class AdvancedDrawingConverterBranchPartialTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_skip_chart_payload_when_sheet_null_or_payload_array() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // null sheet
            AdvancedDrawingConverter.writeSheetCharts(null, mapper.createObjectNode());
            // payload 为数组 → !isObject 命中
            AdvancedDrawingConverter.writeSheetCharts(sh, mapper.createArrayNode());
            // 不抛错即可
            assertThat(sh.getDrawingPatriarch()).isNull();
        }
    }

    @Test
    void should_skip_chart_item_when_null_or_non_object() throws Exception {
        // 覆盖 L145: item == null / !item.isObject() 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            data.set("a", mapper.nullNode());
            data.set("b", mapper.createArrayNode()); // 非 object
            payload.set("data", data);
            ArrayNode order = mapper.createArrayNode();
            order.add("a");
            order.add("b");
            payload.set("order", order);
            AdvancedDrawingConverter.writeSheetCharts(sh, payload);
            // null/非object 都被跳过，不产生 chart
            org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sh.getDrawingPatriarch();
            if (drawing != null) {
                assertThat(drawing.getCharts()).isEmpty();
            }
        }
    }

    @Test
    void should_skip_shapes_when_sheet_null_or_payload_array() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            AdvancedDrawingConverter.writeSheetShapes(null, mapper.createObjectNode());
            AdvancedDrawingConverter.writeSheetShapes(sh, mapper.createArrayNode());
            assertThat(sh.getDrawingPatriarch()).isNull();
        }
    }

    @Test
    void should_skip_shape_item_when_null_or_non_object() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            data.set("a", mapper.nullNode());
            data.set("b", mapper.createArrayNode());
            payload.set("data", data);
            ArrayNode order = mapper.createArrayNode();
            order.add("a");
            order.add("b");
            payload.set("order", order);
            AdvancedDrawingConverter.writeSheetShapes(sh, payload);
        }
    }

    @Test
    void should_skip_sparkline_when_sheet_null_or_payload_array() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            AdvancedDrawingConverter.writeSheetSparkline(null, mapper.createObjectNode());
            AdvancedDrawingConverter.writeSheetSparkline(sh, mapper.createArrayNode());
            assertThat(sh.getCTWorksheet().isSetExtLst()).isFalse();
        }
    }

    @Test
    void should_handle_chart_without_title() throws Exception {
        // 覆盖 L109: chart.getTitleText() == null 分支
        // 通过创建空 chart（rawXml empty）来构造无 title 的 chart
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sh.createDrawingPatriarch();
            drawing.createChart(new org.apache.poi.xssf.usermodel.XSSFClientAnchor(
                    0, 0, 0, 0, 0, 0, 5, 5));
            // 这个 chart 没有 title
            ObjectNode out = AdvancedDrawingConverter.readSheetCharts(sh, mapper, "u", "S");
            // 数据存在但没有 title 字段
            assertThat(out.path("data").size()).isEqualTo(1);
            JsonNode item = out.path("data").fields().next().getValue();
            // title 应该不存在或为空
            assertThat(item.path("title").isMissingNode() || item.path("title").asText().isEmpty()).isTrue();
        }
    }

    @Test
    void should_iterate_data_without_order_using_field_iterator() throws Exception {
        // 覆盖 L347: while it.hasNext()，即 order 为空时走 data.fields() 路径
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode shape = mapper.createObjectNode();
            shape.put("kind", "shape");
            shape.put("shapeType", "rect");
            ObjectNode st = mapper.createObjectNode();
            ObjectNode from = mapper.createObjectNode();
            from.put("column", 0);
            from.put("columnOffset", 0);
            from.put("row", 0);
            from.put("rowOffset", 0);
            ObjectNode to = mapper.createObjectNode();
            to.put("column", 2);
            to.put("columnOffset", 0);
            to.put("row", 2);
            to.put("rowOffset", 0);
            st.set("from", from);
            st.set("to", to);
            shape.set("sheetTransform", st);
            data.set("s1", shape);
            payload.set("data", data);
            // 不设 order，迫使走 data.fields() 路径
            AdvancedDrawingConverter.writeSheetShapes(sh, payload);
            org.apache.poi.xssf.usermodel.XSSFDrawing drawing = sh.getDrawingPatriarch();
            assertThat(drawing.getShapes()).hasSize(1);
        }
    }

    @Test
    void should_handle_simple_shape_without_text() throws Exception {
        // 覆盖 L251: text == null 或 isEmpty 的分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode shape = mapper.createObjectNode();
            shape.put("kind", "shape");
            shape.put("shapeType", "rect");
            // 不设 text
            ObjectNode st = mapper.createObjectNode();
            ObjectNode from = mapper.createObjectNode();
            from.put("column", 0);
            from.put("columnOffset", 0);
            from.put("row", 0);
            from.put("rowOffset", 0);
            ObjectNode to = mapper.createObjectNode();
            to.put("column", 2);
            to.put("columnOffset", 0);
            to.put("row", 2);
            to.put("rowOffset", 0);
            st.set("from", from);
            st.set("to", to);
            shape.set("sheetTransform", st);
            data.set("s1", shape);
            payload.set("data", data);
            AdvancedDrawingConverter.writeSheetShapes(sh, payload);

            // 读回看 text 不存在
            ObjectNode out = AdvancedDrawingConverter.readSheetShapes(sh, mapper, "u", "S");
            JsonNode item = out.path("data").fields().next().getValue();
            assertThat(item.path("text").isMissingNode()).isTrue();
        }
    }
}
