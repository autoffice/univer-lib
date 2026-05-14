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
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Iterator;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 定义名称双向转换：构造 Univer SHEET_DEFINED_NAME_PLUGIN payload → 写入 POI workbook
 * → 回读 → 核对字段。
 */
class DefinedNameConverterTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_when_workbook_has_no_names() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Sheet1");
            ObjectNode out = DefinedNameConverter.readWorkbookDefinedNames(wb, mapper);
            assertThat(out.size()).isEqualTo(0);
        }
    }

    @Test
    void should_roundtrip_workbook_scope_name() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode item = mapper.createObjectNode();
        item.put("id", "dn1");
        item.put("name", "Rate");
        item.put("formulaOrRefString", "Sheet1!$A$1");
        item.put("comment", "tax rate");
        payload.set("dn1", item);

        ObjectNode roundTrip = writeReadBack(payload, "Sheet1");
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode got = roundTrip.fields().next().getValue();
        assertThat(got.path("name").asText()).isEqualTo("Rate");
        assertThat(got.path("formulaOrRefString").asText()).contains("Sheet1");
        assertThat(got.path("comment").asText()).isEqualTo("tax rate");
        // workbook scope: 没有 localSheetId
        assertThat(got.path("localSheetId").isMissingNode()
                || got.path("localSheetId").asText().isEmpty()).isTrue();
    }

    @Test
    void should_roundtrip_sheet_scope_name_and_preserve_localSheetId() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode item = mapper.createObjectNode();
        item.put("id", "dn2");
        item.put("name", "Local");
        item.put("formulaOrRefString", "Sheet1!$B$2");
        item.put("localSheetId", "Sheet1");
        payload.set("dn2", item);

        ObjectNode roundTrip = writeReadBack(payload, "Sheet1");
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode got = roundTrip.fields().next().getValue();
        assertThat(got.path("name").asText()).isEqualTo("Local");
        assertThat(got.path("localSheetId").asText()).isEqualTo("Sheet1");
    }

    @Test
    void should_skip_items_missing_name_or_formula() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode ok = mapper.createObjectNode();
        ok.put("id", "a");
        ok.put("name", "Good");
        ok.put("formulaOrRefString", "Sheet1!$A$1");
        payload.set("a", ok);
        // 缺少 name
        ObjectNode noName = mapper.createObjectNode();
        noName.put("id", "b");
        noName.put("formulaOrRefString", "Sheet1!$B$1");
        payload.set("b", noName);
        // 缺少 formula
        ObjectNode noFormula = mapper.createObjectNode();
        noFormula.put("id", "c");
        noFormula.put("name", "BadNoFormula");
        payload.set("c", noFormula);

        ObjectNode roundTrip = writeReadBack(payload, "Sheet1");
        assertThat(roundTrip.size()).isEqualTo(1);
        JsonNode got = roundTrip.fields().next().getValue();
        assertThat(got.path("name").asText()).isEqualTo("Good");
    }

    @Test
    void should_skip_builtin_names_on_read() throws Exception {
        // 给 workbook 建一条 xlsx 内置名称（以 "_xlnm." 开头），读路径应该过滤掉
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Sheet1");
            XSSFName builtin = wb.createName();
            builtin.setNameName("_xlnm.Print_Area");
            builtin.setRefersToFormula("Sheet1!$A$1:$B$2");
            // 再加一条用户命名
            XSSFName user = wb.createName();
            user.setNameName("MyRange");
            user.setRefersToFormula("Sheet1!$A$1:$A$10");
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            ObjectNode got = DefinedNameConverter.readWorkbookDefinedNames(wb2, mapper);
            assertThat(got.size()).isEqualTo(1);
            Iterator<Map.Entry<String, JsonNode>> it = got.fields();
            Map.Entry<String, JsonNode> only = it.next();
            assertThat(only.getValue().path("name").asText()).isEqualTo("MyRange");
        }
    }

    /** Write payload into a fresh workbook (with given sheet), reopen, read back. */
    private ObjectNode writeReadBack(ObjectNode payload, String sheetName) throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet(sheetName);
            DefinedNameConverter.writeWorkbookDefinedNames(wb, payload);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb2 = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            return DefinedNameConverter.readWorkbookDefinedNames(wb2, mapper);
        }
    }
}
