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

import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * CellConverter 边角分支：错误值、空白带样式、formula 不同缓存类型、字符串 styleId、null 输入等。
 */
class CellConverterBranchTest {

    private XSSFCell newCell(XSSFWorkbook wb) {
        return wb.createSheet().createRow(0).createCell(0);
    }

    @Test
    void should_handle_null_src_silently() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // 不应抛出
            cc.writeCell(c, null);
            assertThat(c.toString()).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_null_for_null_cell_on_read() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            assertThat(cc.readCell(null)).isNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_null_for_blank_cell_without_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // 默认 _NONE/BLANK，没设过任何东西
            assertThat(cc.readCell(c)).isNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_data_with_style_for_blank_cell_with_non_default_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            CellConverter cc = new CellConverter(sc);
            XSSFCell c = newCell(wb);
            // 给单元格挂一个非默认样式但不写值
            XSSFCellStyle style = sc.toPoiStyle(new IStyleData().setFf("Arial").setFs(11));
            c.setCellStyle(style);
            ICellData back = cc.readCell(c);
            assertThat(back).isNotNull();
            assertThat(back.getS()).isInstanceOf(String.class); // styleId
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_strip_leading_equals_from_formula() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setF("=A1+B1").setT(CellValueType.NUMBER).setV(0.0));
            assertThat(c.getCellFormula()).isEqualTo("A1+B1");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_set_cached_boolean_value_for_formula() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setF("ISNUMBER(A1)").setT(CellValueType.BOOLEAN).setV(1));
            assertThat(c.getCellFormula()).isEqualTo("ISNUMBER(A1)");
            assertThat(c.getBooleanCellValue()).isTrue();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_set_cached_string_value_for_formula() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setF("UPPER(A1)").setT(CellValueType.STRING).setV("HELLO"));
            assertThat(c.getCellFormula()).isEqualTo("UPPER(A1)");
            // 缓存值应当是字符串
            assertThat(c.getStringCellValue()).isEqualTo("HELLO");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_cached_value_when_v_is_null() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setF("A1+B1").setT(CellValueType.NUMBER));
            assertThat(c.getCellFormula()).isEqualTo("A1+B1");
            // 没有缓存值，POI 默认 0
            assertThat(c.getNumericCellValue()).isEqualTo(0.0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_write_value_when_t_string_with_number_value() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // 没有显式 t 但 v 是数字 → 自动检测为数字
            cc.writeCell(c, new ICellData().setV(42));
            assertThat(c.getNumericCellValue()).isEqualTo(42.0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_auto_convert_value_when_no_type_and_no_number() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // 没有 t 也不是 Number/String → 走 default 字符串化分支
            cc.writeCell(c, new ICellData().setV(true));
            // Boolean → "true"
            assertThat(c.getStringCellValue()).isEqualTo("true");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_when_value_is_null() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // 没有公式也没有 v → 没有任何写入
            cc.writeCell(c, new ICellData().setT(CellValueType.STRING));
            // POI 默认 _NONE
            assertThat(c.getCellType().toString()).isIn("_NONE", "BLANK");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_convert_string_boolean_value_in_force_text() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // FORCE_TEXT 走 setCellValue(String)
            cc.writeCell(c, new ICellData().setV("123").setT(CellValueType.FORCE_TEXT));
            assertThat(c.getStringCellValue()).isEqualTo("123");
            assertThat(c.getCellStyle().getQuotePrefixed()).isTrue();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_error_cell_as_string() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellErrorValue(FormulaError.DIV0.getCode());
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.STRING);
            // 错误码以数字字符串形式回填
            assertThat(back.getV().toString()).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_formula_with_string_cached_value() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellFormula("UPPER(\"hello\")");
            c.setCellValue("HELLO");
            ICellData back = cc.readCell(c);
            assertThat(back.getF()).isEqualTo("UPPER(\"hello\")");
            assertThat(back.getT()).isEqualTo(CellValueType.STRING);
            assertThat(back.getV()).isEqualTo("HELLO");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_formula_with_boolean_cached_value() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellFormula("TRUE()");
            c.setCellValue(true);
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.BOOLEAN);
            assertThat(((Number) back.getV()).intValue()).isEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_formula_with_error_cached_value() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            c.setCellFormula("1/0");
            c.setCellErrorValue(FormulaError.DIV0.getCode());
            ICellData back = cc.readCell(c);
            assertThat(back.getT()).isEqualTo(CellValueType.STRING);
            assertThat(back.getF()).isEqualTo("1/0");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_string_style_id_on_apply() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            // s 是字符串 styleId 时，CellConverter 不解析（由 WorkbookConverter 解析）
            cc.writeCell(c, new ICellData().setV("x").setT(CellValueType.STRING).setS("abc123"));
            // 不抛错即可
            assertThat(c.getStringCellValue()).isEqualTo("x");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_convert_boolean_value_from_string() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellConverter cc = new CellConverter(new StyleConverter(wb));
            XSSFCell c = newCell(wb);
            cc.writeCell(c, new ICellData().setV("true").setT(CellValueType.BOOLEAN));
            assertThat(c.getBooleanCellValue()).isTrue();

            XSSFCell c2 = wb.getSheetAt(0).createRow(1).createCell(0);
            cc.writeCell(c2, new ICellData().setV(0).setT(CellValueType.BOOLEAN));
            assertThat(c2.getBooleanCellValue()).isFalse();
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
