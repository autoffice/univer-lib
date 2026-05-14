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

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class WorksheetConverterTest {

    private WorksheetConverter newConverter(XSSFWorkbook wb) {
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        return new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());
    }

    @Test
    void should_write_cells_merges_and_read_back() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sheet = wb.createSheet("Sheet1");

            IWorksheetData src = new IWorksheetData().setId("s1").setName("Sheet1");
            Map<Integer, ICellData> row0 = new LinkedHashMap<>();
            row0.put(0, new ICellData().setV("A1").setT(CellValueType.STRING));
            row0.put(1, new ICellData().setV(42.0).setT(CellValueType.NUMBER));
            src.getCellData().put(0, row0);
            src.setMergeData(Arrays.asList(
                new IRange().setStartRow(0).setStartColumn(0).setEndRow(1).setEndColumn(1)));

            wsc.writeSheet(sheet, src);

            assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("A1:B2");
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("A1");

            IWorksheetData back = wsc.readSheet(sheet);
            assertThat(back.getMergeData()).hasSize(1);
            assertThat(back.getCellData().get(0).get(0).getV()).isEqualTo("A1");
            assertThat(((Number) back.getCellData().get(0).get(1).getV()).doubleValue()).isEqualTo(42.0);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void should_roundtrip_freeze_and_gridlines_and_rtl() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sheet = wb.createSheet("Sheet1");

            IWorksheetData src = new IWorksheetData().setId("s1").setName("Sheet1")
                .setFreeze(new IFreeze().setStartRow(1).setStartColumn(1).setXSplit(1).setYSplit(1))
                .setShowGridlines(BooleanNumber.FALSE)
                .setRightToLeft(BooleanNumber.TRUE);

            wsc.writeSheet(sheet, src);

            assertThat(sheet.isDisplayGridlines()).isFalse();
            assertThat(sheet.isRightToLeft()).isTrue();
            assertThat(sheet.getPaneInformation()).isNotNull();
            assertThat(sheet.getPaneInformation().isFreezePane()).isTrue();

            IWorksheetData back = wsc.readSheet(sheet);
            assertThat(back.getShowGridlines()).isEqualTo(BooleanNumber.FALSE);
            assertThat(back.getRightToLeft()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getFreeze()).isNotNull();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void should_roundtrip_row_and_column_sizes() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sheet = wb.createSheet("Sheet1");

            IWorksheetData src = new IWorksheetData().setId("s1").setName("Sheet1");
            src.getRowData().put(0, new IRowData().setH(40.0));
            src.getColumnData().put(0, new IColumnData().setW(140.0));
            src.getColumnData().put(1, new IColumnData().setHd(BooleanNumber.TRUE));
            // Need a cell somewhere so the row exists
            Map<Integer, ICellData> row0 = new LinkedHashMap<>();
            row0.put(0, new ICellData().setV("x").setT(CellValueType.STRING));
            src.getCellData().put(0, row0);

            wsc.writeSheet(sheet, src);

            assertThat(sheet.isColumnHidden(1)).isTrue();
            assertThat(sheet.getColumnWidth(0)).isGreaterThan(0);
            assertThat(sheet.getRow(0).getHeightInPoints()).isGreaterThan(0);

            IWorksheetData back = wsc.readSheet(sheet);
            assertThat(back.getColumnData().get(1).getHd()).isEqualTo(BooleanNumber.TRUE);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void should_handle_formula_with_si() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            SharedFormulaRegistry sfr = new SharedFormulaRegistry();
            StyleConverter sc = new StyleConverter(wb);
            CellConverter cc = new CellConverter(sc);
            RichTextConverter rc = new RichTextConverter(wb);
            WorksheetConverter custom = new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());

            XSSFSheet sheet = wb.createSheet("Sheet1");

            IWorksheetData src = new IWorksheetData().setId("s1").setName("Sheet1");
            Map<Integer, ICellData> row0 = new LinkedHashMap<>();
            row0.put(0, new ICellData().setF("SUM(A1:B1)").setSi("si1").setT(CellValueType.NUMBER).setV(0.0));
            row0.put(1, new ICellData().setF("SUM(A1:B1)").setSi("si1").setT(CellValueType.NUMBER).setV(0.0));
            src.getCellData().put(0, row0);

            custom.writeSheet(sheet, src);
            sfr.applyOnWorkbook(wb);

            assertThat(sheet.getRow(0).getCell(0).getCellFormula()).isEqualTo("SUM(A1:B1)");
            assertThat(sheet.getRow(0).getCell(1).getCellFormula()).isEqualTo("SUM(A1:B1)");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
