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

import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class CellConverterTest {

    private XSSFCell newCell(XSSFWorkbook wb) {
        return wb.createSheet().createRow(0).createCell(0);
    }

    @Test
    void should_write_string_number_boolean() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c1 = newCell(wb);
            cc.writeCell(c1, new ICellData().setV("A1").setT(CellValueType.STRING));
            XSSFCell c2 = wb.getSheetAt(0).createRow(1).createCell(0);
            cc.writeCell(c2, new ICellData().setV(3.14).setT(CellValueType.NUMBER));
            XSSFCell c3 = wb.getSheetAt(0).createRow(2).createCell(0);
            cc.writeCell(c3, new ICellData().setV(1).setT(CellValueType.BOOLEAN));
            assertThat(c1.getStringCellValue()).isEqualTo("A1");
            assertThat(c2.getNumericCellValue()).isEqualTo(3.14);
            assertThat(c3.getBooleanCellValue()).isTrue();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_force_text_set_quote_prefix() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setV("012.0").setT(CellValueType.FORCE_TEXT));
            assertThat(c.getCellStyle().getQuotePrefixed()).isTrue();
            assertThat(c.getStringCellValue()).isEqualTo("012.0");
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.FORCE_TEXT);
            assertThat(back.getV()).isEqualTo("012.0");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_cache_quote_prefix_variant_for_force_text() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            // 三个具有相同（空）样式的 FORCE_TEXT 单元格应共享一个 quotePrefix 样式变体
            // Three FORCE_TEXT cells with same (null) style must share one quote-prefix variant.
            XSSFCell c1 = newCell(wb);
            cc.writeCell(c1, new ICellData().setV("a").setT(CellValueType.FORCE_TEXT));
            XSSFCell c2 = wb.getSheetAt(0).createRow(1).createCell(0);
            cc.writeCell(c2, new ICellData().setV("b").setT(CellValueType.FORCE_TEXT));
            XSSFCell c3 = wb.getSheetAt(0).createRow(2).createCell(0);
            cc.writeCell(c3, new ICellData().setV("c").setT(CellValueType.FORCE_TEXT));

            // POI 默认样式占 1 个，quotePrefix 变体占 1 个，合计 2
            // POI's default style counts as 1, plus our 1 quote-prefix variant = 2.
            assertThat(wb.getNumCellStyles()).isEqualTo(2);
            assertThat(c1.getCellStyle().getQuotePrefixed()).isTrue();
            assertThat(c1.getCellStyle().getIndex()).isEqualTo(c2.getCellStyle().getIndex());
            assertThat(c2.getCellStyle().getIndex()).isEqualTo(c3.getCellStyle().getIndex());
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_write_formula_and_read_back() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setF("SUM(A1:B1)").setV(3.0).setT(CellValueType.NUMBER));
            ICellData back = cc.readCell(c);
            assertThat(back.getF()).isEqualTo("SUM(A1:B1)");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_string_and_number_back() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellValue("hello");
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.STRING);
            assertThat(back.getV()).isEqualTo("hello");

            XSSFCell c2 = wb.getSheetAt(0).createRow(1).createCell(0);
            c2.setCellValue(42.0);
            ICellData back2 = cc.readCell(c2);
            assertThat(back2.getT()).isEqualTo(CellValueType.NUMBER);
            assertThat(((Number) back2.getV()).doubleValue()).isEqualTo(42.0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_boolean_as_0_or_1() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellValue(true);
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.BOOLEAN);
            assertThat(((Number) back.getV()).intValue()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_apply_inline_style_from_IStyleData() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            IStyleData style = new IStyleData().setFf("Arial").setFs(14);
            cc.writeCell(c, new ICellData().setV("x").setT(CellValueType.STRING).setS(style));
            assertThat(c.getCellStyle().getFont().getFontName()).isEqualTo("Arial");
            assertThat(c.getCellStyle().getFont().getFontHeightInPoints()).isEqualTo((short) 14);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_ignore_custom_field_on_cell() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            Map<String, Object> custom = new HashMap<>();
            custom.put("k", "v");
            cc.writeCell(c, new ICellData().setV("x").setCustom(custom).setT(CellValueType.STRING));
            assertThat(c.getStringCellValue()).isEqualTo("x");
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
