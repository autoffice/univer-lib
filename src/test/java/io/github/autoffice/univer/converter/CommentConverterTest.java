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

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 批注双向转换：构造 Univer note JSON → 写入 POI → 回读 → 核对字段。
 * Bidirectional comment/note conversion round-trip.
 */
class CommentConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_roundtrip_single_note() throws Exception {
        ObjectNode rowMap = mapper.createObjectNode();
        ObjectNode colMap = mapper.createObjectNode();
        ObjectNode note = mapper.createObjectNode();
        note.put("note", "Hello world!");
        note.put("width", 160);
        note.put("height", 100);
        note.put("show", false);
        // C3 = row 2 col 2
        colMap.set("2", note);
        rowMap.set("2", colMap);

        ObjectNode roundTrip = writeReadBack(rowMap);

        JsonNode got = roundTrip.path("2").path("2");
        assertThat(got.isMissingNode()).isFalse();
        assertThat(got.path("note").asText()).isEqualTo("Hello world!");
    }

    @Test
    void should_roundtrip_two_notes_on_same_sheet() throws Exception {
        ObjectNode rowMap = mapper.createObjectNode();
        // C3 (row=2,col=2)
        ObjectNode r2 = mapper.createObjectNode();
        r2.set("2", newNote("Hello world!"));
        rowMap.set("2", r2);
        // C8 (row=7,col=2)
        ObjectNode r7 = mapper.createObjectNode();
        r7.set("2", newNote("Hello,world!"));
        rowMap.set("7", r7);

        ObjectNode roundTrip = writeReadBack(rowMap);

        assertThat(roundTrip.path("2").path("2").path("note").asText()).isEqualTo("Hello world!");
        assertThat(roundTrip.path("7").path("2").path("note").asText()).isEqualTo("Hello,world!");
    }

    @Test
    void should_return_empty_when_sheet_has_no_comments() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("Empty");
            ObjectNode out = CommentConverter.readSheetComments(sh, mapper);
            assertThat(out.size()).isEqualTo(0);
        }
    }

    private ObjectNode newNote(String text) {
        ObjectNode n = mapper.createObjectNode();
        n.put("note", text);
        n.put("width", 160);
        n.put("height", 100);
        n.put("show", false);
        return n;
    }

    /** Write → bytes → reopen → read. */
    private ObjectNode writeReadBack(ObjectNode rowMap) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            CommentConverter.writeSheetComments(sh, rowMap);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return CommentConverter.readSheetComments(wb2.getSheetAt(0), mapper);
        }
    }
}
