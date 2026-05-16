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
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * PictureConverter 边角分支：JPEG MIME、空 source、bad base64、order 缺失走 data 顺序、未知 sourceType。
 */
class PictureConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    /** 1×1 JPEG (FFD8) 字节 base64：用于 mime/jpeg 路径。 */
    private static final String JPEG_BASE64 =
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3+iiigD//2Q==";

    @Test
    void should_handle_null_or_invalid_payload() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            PictureConverter.writeSheetPictures(sh, null);
            PictureConverter.writeSheetPictures(sh, mapper.createObjectNode()); // empty
            PictureConverter.writeSheetPictures(sh, mapper.createArrayNode()); // 非 object
            PictureConverter.writeSheetPictures(null, mapper.createObjectNode());
            ObjectNode out = PictureConverter.readSheetPictures(null, mapper, "u", "s");
            assertThat(out.path("data").size()).isEqualTo(0);
        }
    }

    @Test
    void should_skip_picture_with_empty_data() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.set("data", mapper.createObjectNode()); // 空 data
            PictureConverter.writeSheetPictures(sh, payload);
            assertThat(sh.getDrawingPatriarch()).isNull();
        }
    }

    @Test
    void should_handle_picture_without_order_array() throws Exception {
        // 没有 order 时按 data 自身顺序遍历
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            data.set("img1", validPng());
            data.set("img2", validPng());
            payload.set("data", data);
            // 不设 order
            PictureConverter.writeSheetPictures(sh, payload);
            ObjectNode read = PictureConverter.readSheetPictures(sh, mapper, "u", "S");
            assertThat(read.path("data").size()).isEqualTo(2);
        }
    }

    @Test
    void should_decode_jpeg_picture_via_mime() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("drawingType", 0);
            item.put("imageSourceType", "BASE64");
            item.put("source", "data:image/jpeg;base64," + JPEG_BASE64);
            ObjectNode st = mapper.createObjectNode();
            st.set("from", point(1, 0, 1, 0));
            st.set("to", point(3, 0, 3, 0));
            item.set("sheetTransform", st);
            data.set("a", item);
            payload.set("data", data);
            PictureConverter.writeSheetPictures(sh, payload);
            ObjectNode read = PictureConverter.readSheetPictures(sh, mapper, "u", "S");
            assertThat(read.path("data").size()).isEqualTo(1);
            // mime 应是 image/jpeg
            String src = read.path("data").fields().next().getValue().path("source").asText();
            assertThat(src).startsWith("data:image/jpeg;base64,");
        }
    }

    @Test
    void should_skip_picture_with_unknown_sourceType_url() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("drawingType", 0);
            item.put("imageSourceType", "URL");
            item.put("source", "https://example.com/x.png"); // 非 data: 协议
            ObjectNode st = mapper.createObjectNode();
            st.set("from", point(0, 0, 0, 0));
            st.set("to", point(2, 0, 2, 0));
            item.set("sheetTransform", st);
            data.set("a", item);
            payload.set("data", data);
            PictureConverter.writeSheetPictures(sh, payload);
            // 没图片
            ObjectNode read = PictureConverter.readSheetPictures(sh, mapper, "u", "S");
            assertThat(read.path("data").size()).isEqualTo(0);
        }
    }

    @Test
    void should_skip_picture_with_invalid_base64() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("drawingType", 0);
            item.put("imageSourceType", "BASE64");
            item.put("source", "data:image/png;base64,!@#%^&*()");
            ObjectNode st = mapper.createObjectNode();
            st.set("from", point(0, 0, 0, 0));
            st.set("to", point(1, 0, 1, 0));
            item.set("sheetTransform", st);
            data.set("a", item);
            payload.set("data", data);
            PictureConverter.writeSheetPictures(sh, payload);
            ObjectNode read = PictureConverter.readSheetPictures(sh, mapper, "u", "S");
            assertThat(read.path("data").size()).isEqualTo(0);
        }
    }

    @Test
    void should_skip_picture_with_empty_source() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode data = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("drawingType", 0);
            item.put("imageSourceType", "BASE64");
            item.put("source", "");
            data.set("a", item);
            payload.set("data", data);
            PictureConverter.writeSheetPictures(sh, payload);
            assertThat(sh.getDrawingPatriarch()).isNotNull(); // createDrawingPatriarch 已被调用
        }
    }

    private ObjectNode validPng() {
        ObjectNode item = mapper.createObjectNode();
        item.put("drawingType", 0);
        item.put("imageSourceType", "BASE64");
        item.put("source",
                "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");
        ObjectNode st = mapper.createObjectNode();
        st.set("from", point(0, 0, 0, 0));
        st.set("to", point(1, 0, 1, 0));
        item.set("sheetTransform", st);
        return item;
    }

    private ObjectNode point(int col, int colOff, int row, int rowOff) {
        ObjectNode n = mapper.createObjectNode();
        n.put("column", col);
        n.put("columnOffset", colOff);
        n.put("row", row);
        n.put("rowOffset", rowOff);
        return n;
    }
}
