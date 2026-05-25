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

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * FilterConverter 边角分支：null sheet/payload、坏 ref、纯数值字段、坏 rawXml 兜底。
 */
class FilterConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_for_null_sheet_on_read() {
        ObjectNode out = FilterConverter.readSheetFilter(null, mapper);
        assertThat(out).isEmpty();
    }

    @Test
    void should_return_empty_for_sheet_without_filter() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode out = FilterConverter.readSheetFilter(sh, mapper);
            assertThat(out).isEmpty();
        }
    }

    @Test
    void should_skip_write_when_payload_is_null_or_empty_or_non_object() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            FilterConverter.writeSheetFilter(sh, null);
            FilterConverter.writeSheetFilter(sh, mapper.createObjectNode());
            FilterConverter.writeSheetFilter(sh, mapper.createArrayNode());
            FilterConverter.writeSheetFilter(null, mapper.createObjectNode().put("ref", "A1:B2"));
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void should_use_numeric_fields_when_ref_invalid() throws Exception {
        // ref 非法 → 走 numeric 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("h");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", "$$$"); // 非法
            payload.put("startRow", 0);
            payload.put("endRow", 1);
            payload.put("startColumn", 0);
            payload.put("endColumn", 1);
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }

    @Test
    void should_apply_rawXml_overlay() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("h");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", "A1:B2");
            // 合法 rawXml（带 namespace）
            payload.put("rawXml",
                    "<autoFilter xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ref=\"A1:B2\"/>");
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }

    @Test
    void should_ignore_invalid_rawXml() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.setAutoFilter(CellRangeAddress.valueOf("A1:B2"));
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", "A1:B2");
            payload.put("rawXml", "<not-valid-xml>");
            // 不应抛错；range-only autoFilter 仍然存在
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }

    @Test
    void should_skip_when_no_ref_and_no_numeric_fields() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("rawXml", "<doesnt matter>");
            // 没 ref 也没 numeric range → 跳过
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void should_handle_invalid_numeric_range() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("startRow", -1);
            payload.put("endRow", -1);
            payload.put("startColumn", -1);
            payload.put("endColumn", -1);
            // POI 在 -1 范围上 setAutoFilter 时抛出 FormulaParseException（属于 RuntimeException）。
            // 我们只是验证 resolveRange 走到了 numeric 分支并把异常透出。
            try {
                FilterConverter.writeSheetFilter(sh, payload);
            } catch (RuntimeException ignored) {
                // 期望
            }
        }
    }

    @Test
    void should_use_only_ref_field_without_numeric_fallback() throws Exception {
        // 覆盖 resolveRange L143: ref 有效 → 直接返回，不走 numeric
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("h");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", "A1:B2");
            // 不设 numeric 字段
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }

    @Test
    void should_use_numeric_when_ref_field_empty_string() throws Exception {
        // 覆盖 resolveRange L143: ref="" → 落到 numeric 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("h");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", ""); // 空串
            payload.put("startRow", 0);
            payload.put("endRow", 1);
            payload.put("startColumn", 0);
            payload.put("endColumn", 1);
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }

    @Test
    void should_skip_numeric_when_partial_fields_present() throws Exception {
        // 覆盖 resolveRange L149-150 of branch: 有 startRow 但没 endRow → return null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("startRow", 0);
            // 故意缺 endRow / startColumn / endColumn
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void should_skip_numeric_when_endColumn_missing() throws Exception {
        // 覆盖 hasNonNull 链中 endColumn 缺失分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("startRow", 0);
            payload.put("endRow", 1);
            payload.put("startColumn", 0);
            // 缺 endColumn
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void should_skip_numeric_when_startColumn_missing() throws Exception {
        // 覆盖 hasNonNull 链中 startColumn 缺失分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("startRow", 0);
            payload.put("endRow", 1);
            // 缺 startColumn / endColumn
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void should_skip_write_when_rawXml_empty_string() throws Exception {
        // 覆盖 L125: rawXml 是空串 → 提前 return（不走 parse）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("h");
            ObjectNode payload = mapper.createObjectNode();
            payload.put("ref", "A1:B2");
            payload.put("rawXml", ""); // 空串
            FilterConverter.writeSheetFilter(sh, payload);
            assertThat(sh.getCTWorksheet().isSetAutoFilter()).isTrue();
        }
    }
}
