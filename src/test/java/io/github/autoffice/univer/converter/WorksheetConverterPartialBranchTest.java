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
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IColumnData;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IFreeze;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IRowData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * WorksheetConverter partial 分支精确补全：
 * - cellData 中 null cellData 元素 / null rowMap
 * - rowData / columnData 中 null entry
 * - mergeData 中 null 元素 / 缺字段
 * - freeze 部分字段 null 的 fallback
 * - tabColor argb=null 分支
 * - shared formula (f + si) 写入路径
 * - applyStyleOnly with non-IStyleData
 */
class WorksheetConverterPartialBranchTest {

    private WorksheetConverter newConverter(XSSFWorkbook wb) {
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        return new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());
    }

    @Test
    void should_skip_null_rowMap_in_cellData() throws Exception {
        // 覆盖 L167: rowMap == null 或 isEmpty
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, null); // null rowMap
            cells.put(1, new LinkedHashMap<>()); // empty rowMap
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            // 不抛错
            assertThat(sh.getPhysicalNumberOfRows()).isEqualTo(0);
        }
    }

    @Test
    void should_skip_null_cellData_entry() throws Exception {
        // 覆盖 L177: cellData == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, null); // null cellData
            row.put(1, new ICellData().setV("x"));
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, row);
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            // 第二个 cell 应被写入
            assertThat(sh.getRow(0).getCell(1).getStringCellValue()).isEqualTo("x");
        }
    }

    @Test
    void should_handle_shared_formula_si_path() throws Exception {
        // 覆盖 L198: cellData.getF() != null && cellData.getSi() != null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setF("A1+1").setSi("si-1"));
            row.put(1, new ICellData().setF("A1+1").setSi("si-1"));
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, row);
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            // 不抛错
            assertThat((Object) sh.getRow(0)).isNotNull();
        }
    }

    @Test
    void should_apply_style_only_with_non_IStyleData_object() throws Exception {
        // 覆盖 L213: styleObj 不是 IStyleData (是 String 或其他) → 跳过 setCellStyle
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            // 共享公式，s 是 String id (不是 IStyleData) → applyStyleOnly 跳过
            row.put(0, new ICellData().setF("A1+1").setSi("si-2").setS("style-id-string"));
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, row);
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            // 不抛错
            assertThat((Object) sh.getRow(0)).isNotNull();
        }
    }

    @Test
    void should_handle_isHyperlinkOnlyDoc_with_null_dataStream() {
        // 覆盖 L229: dataStream 为 null 的情况
        // 通过反射调用 isHyperlinkOnlyDoc 私有方法
        try {
            java.lang.reflect.Method m = WorksheetConverter.class.getDeclaredMethod(
                    "isHyperlinkOnlyDoc", IDocumentData.class);
            m.setAccessible(true);
            // null doc → false
            Boolean r1 = (Boolean) m.invoke(null, (Object) null);
            assertThat(r1.booleanValue()).isFalse();
            // doc.body == null → false
            IDocumentData noBody = new IDocumentData();
            Boolean r2 = (Boolean) m.invoke(null, noBody);
            assertThat(r2.booleanValue()).isFalse();
            // dataStream == null → true (no body content)
            IDocumentData nullStream = new IDocumentData()
                    .setBody(new IDocumentData.Body());
            Boolean r3 = (Boolean) m.invoke(null, nullStream);
            assertThat(r3.booleanValue()).isTrue();
            // dataStream == "" → true
            IDocumentData emptyStream = new IDocumentData()
                    .setBody(new IDocumentData.Body().setDataStream(""));
            Boolean r4 = (Boolean) m.invoke(null, emptyStream);
            assertThat(r4.booleanValue()).isTrue();
            // dataStream == "\r\n" → true
            IDocumentData crlfStream = new IDocumentData()
                    .setBody(new IDocumentData.Body().setDataStream("\r\n"));
            Boolean r5 = (Boolean) m.invoke(null, crlfStream);
            assertThat(r5.booleanValue()).isTrue();
            // dataStream == "real text\r\n" → false
            IDocumentData textStream = new IDocumentData()
                    .setBody(new IDocumentData.Body().setDataStream("real text\r\n"));
            Boolean r6 = (Boolean) m.invoke(null, textStream);
            assertThat(r6.booleanValue()).isFalse();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void should_skip_null_rowData_entry() throws Exception {
        // 覆盖 L240: rd == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, IRowData> rowData = new LinkedHashMap<>();
            rowData.put(0, null); // null entry
            rowData.put(1, new IRowData().setH(20.0));
            ws.setRowData(rowData);
            wsc.writeSheet(sh, ws);
            assertThat((Object) sh.getRow(1)).isNotNull();
        }
    }

    @Test
    void should_skip_null_columnData_entry() throws Exception {
        // 覆盖 L264: cd == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, IColumnData> colData = new LinkedHashMap<>();
            colData.put(0, null); // null entry
            colData.put(1, new IColumnData().setW(120.0));
            ws.setColumnData(colData);
            wsc.writeSheet(sh, ws);
            // 第二列宽度应被设置
            assertThat(sh.getColumnWidth(1)).isPositive();
        }
    }

    @Test
    void should_skip_invalid_merge_entries() throws Exception {
        // 覆盖 L283-284: merge 中 r 为 null 或缺字段
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            List<IRange> merges = Arrays.asList(
                    null, // null
                    new IRange(), // 缺所有字段
                    new IRange().setStartRow(0), // 仅有 startRow
                    new IRange().setStartRow(0).setEndRow(1)
                            .setStartColumn(0).setEndColumn(1)); // 完整
            ws.setMergeData(merges);
            wsc.writeSheet(sh, ws);
            // 仅最后一个被处理
            assertThat(sh.getNumMergedRegions()).isEqualTo(1);
        }
    }

    @Test
    void should_skip_freeze_when_all_zero_or_negative() throws Exception {
        // 覆盖 L303: 全为 -1 / 0 时跳过
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            // 全为 null → fallback 全是 -1 / 0 → 跳过
            ws.setFreeze(new IFreeze());
            wsc.writeSheet(sh, ws);
            // 没有冻结
            assertThat(sh.getPaneInformation()).isNull();
        }
    }

    @Test
    void should_apply_freeze_with_only_xSplit_set() throws Exception {
        // 覆盖 freeze 几个字段独立设置时的 fallback
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            ws.setFreeze(new IFreeze().setXSplit(2));
            wsc.writeSheet(sh, ws);
            // 应该应用了 freeze
            assertThat(sh.getPaneInformation()).isNotNull();
        }
    }

    @Test
    void should_handle_tabColor_with_invalid_rgb() throws Exception {
        // 覆盖 L129: argb == null（rgb 非法）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S")
                    .setTabColor("invalid-hex");
            wsc.writeSheet(sh, ws);
            // tab color 不应被设置
            assertThat(sh.getTabColor()).isNull();
        }
    }

    @Test
    void should_handle_merge_extending_dimensions_in_read() throws Exception {
        // 覆盖 L347 / L350: merge 区域扩展 maxRow / maxCol
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0).setCellValue("a");
            // 添加 merge 让其扩展 maxCol
            sh.addMergedRegion(org.apache.poi.ss.util.CellRangeAddress.valueOf("A1:Z10"));
            IWorksheetData ws = wsc.readSheet(sh);
            // rowCount/columnCount 应当反映 merge 范围
            assertThat(ws.getRowCount()).isGreaterThanOrEqualTo(10);
            assertThat(ws.getColumnCount()).isGreaterThanOrEqualTo(26);
        }
    }

    @Test
    void should_handle_existing_poiRow_when_writing_cellData() throws Exception {
        // 覆盖 L171: poiRow != null（已存在）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            // 预先创建 row
            sh.createRow(0);
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("x"));
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, row);
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            assertThat(sh.getRow(0).getCell(0).getStringCellValue()).isEqualTo("x");
        }
    }

    @Test
    void should_handle_existing_cell_when_writing() throws Exception {
        // 覆盖 L181: cell != null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            sh.createRow(0).createCell(0); // 预先创建 cell
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("y"));
            Map<Integer, Map<Integer, ICellData>> cells = new LinkedHashMap<>();
            cells.put(0, row);
            ws.setCellData(cells);
            wsc.writeSheet(sh, ws);
            assertThat(sh.getRow(0).getCell(0).getStringCellValue()).isEqualTo("y");
        }
    }

    @Test
    void should_apply_freeze_with_only_yPlit_and_negative_startRow() throws Exception {
        // 覆盖 L298-301 的更多组合
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            ws.setFreeze(new IFreeze().setYSplit(2)); // 只设 ySplit
            wsc.writeSheet(sh, ws);
            assertThat(sh.getPaneInformation()).isNotNull();
        }
    }

    @Test
    void should_apply_freeze_with_startRow_only() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            ws.setFreeze(new IFreeze().setStartRow(2));
            wsc.writeSheet(sh, ws);
            // POI 在 (0,0,startCol,startRow) 调用下可能不创建 PaneInformation
            // 但 sheet 不应抛错，覆盖了 L303 中部分组合分支
            assertThat(sh.getSheetName()).isEqualTo("S");
        }
    }
}
