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
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Base64;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 图片双向转换：构造 Univer ISheetImage JSON → 写入 POI → 回读 → 核对字段。
 * Bidirectional image round-trip.
 */
class PictureConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    /** 一张 1x1 透明 PNG（合法字节，能被 POI 识别）。 */
    private static final String BASE64_1PX_PNG =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    @Test
    void should_roundtrip_single_png() throws Exception {
        ObjectNode payload = buildPayload("img-1",
                "data:image/png;base64," + BASE64_1PX_PNG,
                1, 10, 3, 20, 5, 30, 7, 40);

        ObjectNode roundTrip = writeReadBack(payload);

        JsonNode data = roundTrip.path("data");
        assertThat(data.size()).isEqualTo(1);
        JsonNode item = data.fields().next().getValue();
        assertThat(item.path("drawingType").asInt()).isEqualTo(0);
        assertThat(item.path("imageSourceType").asText()).isEqualTo("BASE64");
        assertThat(item.path("source").asText()).startsWith("data:image/png;base64,");

        JsonNode from = item.path("sheetTransform").path("from");
        JsonNode to = item.path("sheetTransform").path("to");
        assertThat(from.path("column").asInt()).isEqualTo(1);
        assertThat(from.path("row").asInt()).isEqualTo(3);
        assertThat(to.path("column").asInt()).isEqualTo(5);
        assertThat(to.path("row").asInt()).isEqualTo(7);
        // columnOffset / rowOffset 是像素；POI 在 round-trip 中保留 EMU，所以 ±1 像素误差接受
        assertThat(from.path("columnOffset").asInt()).isEqualTo(10);
        assertThat(from.path("rowOffset").asInt()).isEqualTo(20);
        assertThat(to.path("columnOffset").asInt()).isEqualTo(30);
        assertThat(to.path("rowOffset").asInt()).isEqualTo(40);

        // 回读的 BASE64 内容必须能解码且与原始相同
        String src = item.path("source").asText();
        String base64 = src.substring(src.indexOf(',') + 1);
        byte[] rt = Base64.getDecoder().decode(base64);
        byte[] orig = Base64.getDecoder().decode(BASE64_1PX_PNG);
        assertThat(rt).isEqualTo(orig);
    }

    @Test
    void should_roundtrip_two_pictures_on_same_sheet() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        data.set("a", imageItem("a",
                "data:image/png;base64," + BASE64_1PX_PNG, 0, 0, 0, 0, 2, 0, 2, 0));
        data.set("b", imageItem("b",
                "data:image/png;base64," + BASE64_1PX_PNG, 5, 0, 5, 0, 7, 0, 7, 0));
        payload.set("data", data);
        com.fasterxml.jackson.databind.node.ArrayNode order = mapper.createArrayNode();
        order.add("a");
        order.add("b");
        payload.set("order", order);

        ObjectNode roundTrip = writeReadBack(payload);
        assertThat(roundTrip.path("data").size()).isEqualTo(2);
    }

    @Test
    void should_skip_non_image_drawing_type() throws Exception {
        ObjectNode payload = buildPayload("shape-1",
                "data:image/png;base64," + BASE64_1PX_PNG,
                1, 0, 1, 0, 2, 0, 2, 0);
        // 把 drawingType 改成 shape（1），应该被 writer 忽略
        ((ObjectNode) payload.path("data").path("shape-1")).put("drawingType", 1);

        ObjectNode roundTrip = writeReadBack(payload);
        assertThat(roundTrip.path("data").size()).isEqualTo(0);
    }

    @Test
    void should_return_empty_when_sheet_has_no_drawing() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("Empty");
            ObjectNode out = PictureConverter.readSheetPictures(sh, mapper, "unit-1", "Empty");
            assertThat(out.path("data").size()).isEqualTo(0);
            assertThat(out.path("order").size()).isEqualTo(0);
        }
    }

    private ObjectNode buildPayload(String id, String source,
                                    int fromCol, int fromColOff, int fromRow, int fromRowOff,
                                    int toCol, int toColOff, int toRow, int toRowOff) {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        data.set(id, imageItem(id, source,
                fromCol, fromColOff, fromRow, fromRowOff,
                toCol, toColOff, toRow, toRowOff));
        payload.set("data", data);
        com.fasterxml.jackson.databind.node.ArrayNode order = mapper.createArrayNode();
        order.add(id);
        payload.set("order", order);
        return payload;
    }

    private ObjectNode imageItem(String id, String source,
                                 int fromCol, int fromColOff, int fromRow, int fromRowOff,
                                 int toCol, int toColOff, int toRow, int toRowOff) {
        ObjectNode item = mapper.createObjectNode();
        item.put("unitId", "unit-1");
        item.put("subUnitId", "S1");
        item.put("drawingId", id);
        item.put("drawingType", 0);
        item.put("imageSourceType", "BASE64");
        item.put("source", source);
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
        item.set("sheetTransform", st);
        return item;
    }

    private ObjectNode writeReadBack(ObjectNode payload) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            PictureConverter.writeSheetPictures(sh, payload);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return PictureConverter.readSheetPictures(wb2.getSheetAt(0), mapper, "unit-1", "S1");
        }
    }
}
