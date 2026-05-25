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

import static org.assertj.core.api.Assertions.assertThat;

/**
 * DefinedNameConverter 边角分支：null inputs / 同名替换 / 不存在的 sheet localSheetId / 非 object 节点。
 */
class DefinedNameConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_for_null_workbook_or_no_names() {
        ObjectNode out1 = DefinedNameConverter.readWorkbookDefinedNames(null, mapper);
        assertThat(out1.size()).isEqualTo(0);
    }

    @Test
    void should_skip_write_for_null_or_empty_inputs() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S");
            DefinedNameConverter.writeWorkbookDefinedNames(null, mapper.createObjectNode());
            DefinedNameConverter.writeWorkbookDefinedNames(wb, null);
            DefinedNameConverter.writeWorkbookDefinedNames(wb, mapper.createArrayNode()); // 非 object
            DefinedNameConverter.writeWorkbookDefinedNames(wb, mapper.createObjectNode()); // 空
            assertThat(wb.getAllNames()).isEmpty();
        }
    }

    @Test
    void should_skip_non_object_item_in_payload() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.set("a", mapper.nullNode());
            payload.set("b", mapper.createArrayNode()); // 非 object
            DefinedNameConverter.writeWorkbookDefinedNames(wb, payload);
            assertThat(wb.getAllNames()).isEmpty();
        }
    }

    @Test
    void should_replace_existing_name_when_collision() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S");
            // 先创建一条同名
            XSSFName existing = wb.createName();
            existing.setNameName("Same");
            existing.setRefersToFormula("S!$A$1");

            // 写入同名的不同公式 → 应替换
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("id", "x");
            item.put("name", "Same");
            item.put("formulaOrRefString", "S!$B$2");
            payload.set("x", item);

            DefinedNameConverter.writeWorkbookDefinedNames(wb, payload);
            assertThat(wb.getAllNames()).hasSize(1);
            assertThat(wb.getName("Same").getRefersToFormula()).contains("$B$2");
        }
    }

    @Test
    void should_skip_when_localSheetId_not_found() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            ObjectNode item = mapper.createObjectNode();
            item.put("id", "x");
            item.put("name", "X");
            item.put("formulaOrRefString", "S!$A$1");
            item.put("localSheetId", "GhostSheet"); // 不存在
            payload.set("x", item);

            DefinedNameConverter.writeWorkbookDefinedNames(wb, payload);
            // 仍然创建（workbook scope）
            assertThat(wb.getAllNames()).hasSize(1);
        }
    }

    @Test
    void should_read_hidden_name_flag() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S");
            XSSFName n = wb.createName();
            n.setNameName("Hidden");
            n.setRefersToFormula("S!$A$1");
            n.setComment("note");
            // 通过反射访问 protected getCTName 设 hidden
            java.lang.reflect.Method m = XSSFName.class.getDeclaredMethod("getCTName");
            m.setAccessible(true);
            Object ct = m.invoke(n);
            ct.getClass().getMethod("setHidden", boolean.class).invoke(ct, true);

            ObjectNode out = DefinedNameConverter.readWorkbookDefinedNames(wb, mapper);
            assertThat(out.size()).isEqualTo(1);
            JsonNode item = out.fields().next().getValue();
            assertThat(item.path("hidden").asBoolean()).isTrue();
            assertThat(item.path("comment").asText()).isEqualTo("note");
        }
    }
}
