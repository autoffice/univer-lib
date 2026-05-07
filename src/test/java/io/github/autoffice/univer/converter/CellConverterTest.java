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
