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
package io.github.autoffice.univer.io;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

/**
 * 一次性工具类：在现有 src/test/resources/test.xlsx 基础上注入 9 种原生 XLSX 特性，
 * 方便 end-to-end 验证 lib 的转换覆盖面。执行方式：
 * <pre>
 *   ./mvnw -q test-compile
 *   java -cp target/test-classes:target/classes:$(./mvnw -q dependency:build-classpath -Dmdep.outputFile=/tmp/cp && cat /tmp/cp) \
 *        io.github.autoffice.univer.io.TestXlsxFixtureBuilder
 * </pre>
 *
 * <p>保留原 sheet 结构与已有断言依赖（Cell.A6 富文本、Conditional Format sheet 名称等），
 * 只新增以下原生 XLSX 构件：
 * <ul>
 *   <li>Cell sheet：hyperlink（URL / 工作表内部）</li>
 *   <li>workbook：全局 + sheet-scope defined name</li>
 *   <li>Data Verification sheet：下拉列表 + 数字范围 data validation</li>
 *   <li>Table sheet：auto-filter + XSSFTable（包含样式）</li>
 *   <li>Picture sheet：textbox shape + connector shape</li>
 * </ul>
 * Chart / Pivot / Sparkline 已有对应 sheet（fixture 已含 chart，pivot 由独立测试构造），
 * 这里不再重复生成，避免和现有断言冲突。
 */
public final class TestXlsxFixtureBuilder {

    private TestXlsxFixtureBuilder() {}

    public static void main(String[] args) throws IOException {
        Path src = Paths.get("src/test/resources/test.xlsx");
        Path tmp = Files.createTempFile("test-fixture-", ".xlsx");
        Files.copy(src, tmp, StandardCopyOption.REPLACE_EXISTING);

        try (FileInputStream in = new FileInputStream(tmp.toFile())) {
            XSSFWorkbook wb = new XSSFWorkbook(in);
            injectHyperlinks(wb);
            injectDefinedNames(wb);
            injectDataValidations(wb);
            injectTableAndAutoFilter(wb);
            injectShapes(wb);
            injectChart(wb);
            injectPivotTable(wb);
            injectSparkline(wb);
            injectSparkline(wb);
            try (FileOutputStream out = new FileOutputStream(src.toFile())) {
                wb.write(out);
            }
            wb.close();
        }
        Files.deleteIfExists(tmp);
        System.out.println("[fixture] rewritten " + src.toAbsolutePath());
    }

    // ------------------------------------------------------------
    // 1. Hyperlinks on Cell sheet
    // ------------------------------------------------------------
    private static void injectHyperlinks(XSSFWorkbook wb) {
        XSSFSheet cell = wb.getSheet("Cell");
        if (cell == null) {
            return;
        }
        CreationHelper helper = wb.getCreationHelper();
        // 外部 URL：C13（保留已有单元格内容，附加 hyperlink）
        XSSFRow row = ensureRow(cell, 12);
        XSSFCell c13 = ensureCell(row, 2);
        c13.setCellValue("Univer Docs");
        Hyperlink url = helper.createHyperlink(HyperlinkType.URL);
        url.setAddress("https://univer.ai");
        c13.setHyperlink(url);

        // 工作表内部：D13 -> Formula sheet
        XSSFCell d13 = ensureCell(row, 3);
        d13.setCellValue("Go to Formula");
        Hyperlink doc = helper.createHyperlink(HyperlinkType.DOCUMENT);
        doc.setAddress("Formula!A1");
        d13.setHyperlink(doc);
    }

    // ------------------------------------------------------------
    // 2. Defined names
    // ------------------------------------------------------------
    private static void injectDefinedNames(XSSFWorkbook wb) {
        // workbook-scope
        XSSFName taxRate = (XSSFName) wb.createName();
        taxRate.setNameName("TaxRate");
        taxRate.setRefersToFormula("0.17");
        taxRate.setComment("Default tax rate used across sheets.");

        // sheet-scope (Formula 工作表)
        int formulaIdx = wb.getSheetIndex("Formula");
        if (formulaIdx >= 0) {
            XSSFName localRange = (XSSFName) wb.createName();
            localRange.setNameName("LocalRange");
            localRange.setRefersToFormula("Formula!$A$1:$B$5");
            localRange.setSheetIndex(formulaIdx);
        }
    }

    // ------------------------------------------------------------
    // 3. Data validations on Data Verification sheet
    // ------------------------------------------------------------
    private static void injectDataValidations(XSSFWorkbook wb) {
        XSSFSheet dv = wb.getSheet("Data Verification");
        if (dv == null) {
            return;
        }
        DataValidationHelper helper = dv.getDataValidationHelper();
        // 下拉列表：A1:A10
        DataValidationConstraint listConstraint = helper.createExplicitListConstraint(
                new String[] {"Yes", "No", "Maybe"});
        CellRangeAddressList listRange = new CellRangeAddressList(0, 9, 0, 0);
        DataValidation listDV = helper.createValidation(listConstraint, listRange);
        listDV.setShowErrorBox(true);
        listDV.createErrorBox("Invalid", "Please choose Yes / No / Maybe");
        dv.addValidationData(listDV);

        // 数字范围：B1:B10 between 1 and 100
        DataValidationConstraint numConstraint = helper.createNumericConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.BETWEEN,
                "1", "100");
        CellRangeAddressList numRange = new CellRangeAddressList(0, 9, 1, 1);
        DataValidation numDV = helper.createValidation(numConstraint, numRange);
        numDV.setShowPromptBox(true);
        numDV.createPromptBox("Range", "Enter integer 1..100");
        dv.addValidationData(numDV);
    }

    // ------------------------------------------------------------
    // 4. Auto-filter + XSSFTable on Table sheet
    // ------------------------------------------------------------
    private static void injectTableAndAutoFilter(XSSFWorkbook wb) {
        XSSFSheet table = wb.getSheet("Table");
        if (table == null) {
            return;
        }
        // 确保 4x3 数据存在：Name/Qty/Price header + 3 data rows
        writeRowAt(table, 0, 0, "Name", "Qty", "Price");
        writeRowAt(table, 1, 0, "Apple", 10, 1.5);
        writeRowAt(table, 2, 0, "Banana", 20, 0.8);
        writeRowAt(table, 3, 0, "Cherry", 5, 3.2);

        // Auto-filter
        AutoFilter af = table.setAutoFilter(new CellRangeAddress(0, 3, 0, 2));
        // 用一下返回，避免 unused warning
        if (af == null) {
            return;
        }

        // XSSFTable (ListObject)，放在和 auto-filter 不重叠的区域：E1:G4
        writeRowAt(table, 0, 4, "Key", "Value", "Note");
        writeRowAt(table, 1, 4, "A", 1.0, "first");
        writeRowAt(table, 2, 4, "B", 2.0, "second");
        writeRowAt(table, 3, 4, "C", 3.0, "third");
        AreaReference area = wb.getCreationHelper().createAreaReference(
                new CellReference(0, 4), new CellReference(3, 6));
        // 避免重复 table
        if (!table.getTables().isEmpty()) {
            return;
        }
        XSSFTable t = table.createTable(area);
        t.setName("DemoTable");
        t.setDisplayName("DemoTable");
        if (t.getColumns().size() == 0) {
            t.createColumn("Key");
            t.createColumn("Value");
            t.createColumn("Note");
        }
        XSSFTableStyleInfo style = (XSSFTableStyleInfo) t.getStyle();
        if (style != null) {
            style.setName("TableStyleMedium9");
            style.setShowRowStripes(true);
        }
    }

    // ------------------------------------------------------------
    // 5. Non-image shapes (textbox + connector) on Picture sheet
    // ------------------------------------------------------------
    private static void injectShapes(XSSFWorkbook wb) {
        XSSFSheet pic = wb.getSheet("Picture");
        if (pic == null) {
            return;
        }
        XSSFDrawing drawing = pic.createDrawingPatriarch();
        // TextBox（POI 5.2.5 没有 TEXT_BOX 常量，文本框语义用 RECT；XSSFSimpleShape 自带 setText 渲染）
        XSSFClientAnchor tbAnchor = new XSSFClientAnchor(0, 0, 0, 0, 1, 15, 4, 18);
        XSSFSimpleShape tb = drawing.createSimpleShape(tbAnchor);
        tb.setText("Hello shape!");
        tb.setShapeType(org.apache.poi.ss.usermodel.ShapeTypes.RECT);

        // Connector
        XSSFClientAnchor connAnchor = new XSSFClientAnchor(0, 0, 0, 0, 6, 15, 9, 18);
        XSSFConnector connector = drawing.createConnector(connAnchor);
        connector.setShapeType(org.apache.poi.ss.usermodel.ShapeTypes.STRAIGHT_CONNECTOR_1);
    }

    // ------------------------------------------------------------
    // 6. Chart on Chart sheet (empty CTChartSpace placeholder)
    // ------------------------------------------------------------
    private static void injectChart(XSSFWorkbook wb) {
        XSSFSheet chart = wb.getSheet("Chart");
        if (chart == null) {
            return;
        }
        // 已存在 chart 时跳过，避免重复
        XSSFDrawing existing = chart.getDrawingPatriarch();
        if (existing != null) {
            for (org.apache.poi.xssf.usermodel.XSSFShape s : existing.getShapes()) {
                if (s instanceof org.apache.poi.xssf.usermodel.XSSFGraphicFrame) {
                    return;
                }
            }
        }
        // 给一些数据让图表有意义
        writeRowAt(chart, 0, 0, "Region", "Value");
        writeRowAt(chart, 1, 0, "East", 10);
        writeRowAt(chart, 2, 0, "West", 20);
        writeRowAt(chart, 3, 0, "North", 15);

        XSSFDrawing drawing = chart.createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 4, 0, 12, 18);
        // 创建空 chart 容器；CTChartSpace 默认即可，AdvancedDrawingConverter 通过 rawXml 保留
        drawing.createChart(anchor);
    }

    // ------------------------------------------------------------
    // 7. Pivot table on PivotTable sheet, source from PivotTableData
    // ------------------------------------------------------------
    private static void injectPivotTable(XSSFWorkbook wb) {
        XSSFSheet pivot = wb.getSheet("PivotTable");
        XSSFSheet source = wb.getSheet("PivotTableData");
        if (pivot == null || source == null) {
            return;
        }
        if (!pivot.getPivotTables().isEmpty()) {
            return;
        }
        // 确保 source 有 header + rows
        writeRowAt(source, 0, 0, "Region", "Amount");
        writeRowAt(source, 1, 0, "East", 10);
        writeRowAt(source, 2, 0, "West", 20);
        writeRowAt(source, 3, 0, "North", 15);

        AreaReference area = wb.getCreationHelper().createAreaReference(
                new CellReference("PivotTableData!A1"),
                new CellReference("PivotTableData!B4"));
        pivot.createPivotTable(area, new CellReference(0, 0), source);
    }

    // ------------------------------------------------------------
    // 8. Sparkline on Sparkline sheet (raw extLst XML)
    // ------------------------------------------------------------
    private static void injectSparkline(XSSFWorkbook wb) {
        XSSFSheet sl = wb.getSheet("Sparkline");
        if (sl == null) {
            return;
        }
        // 已有 extLst 则跳过，避免覆盖
        if (sl.getCTWorksheet().isSetExtLst()) {
            return;
        }
        // 给 sparkline 提供数据：A1:A5 数据 + B1 sparkline 容器
        writeRowAt(sl, 0, 0, 1);
        writeRowAt(sl, 1, 0, 3);
        writeRowAt(sl, 2, 0, 2);
        writeRowAt(sl, 3, 0, 5);
        writeRowAt(sl, 4, 0, 4);
        // 直接拼装最小 sparklineGroups extLst，AdvancedDrawingConverter 会原样保留 rawXml
        String xml = "<extLst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
                + "uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\">"
                + "<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
                + "<x14:sparklineGroup type=\"line\">"
                + "<x14:sparklines>"
                + "<x14:sparkline><xm:f>Sparkline!A1:A5</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline>"
                + "</x14:sparklines>"
                + "</x14:sparklineGroup>"
                + "</x14:sparklineGroups>"
                + "</ext>"
                + "</extLst>";
        try {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList ext =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList.Factory.parse(xml);
            sl.getCTWorksheet().setExtLst(ext);
        } catch (Exception ignored) {
            // best-effort：解析失败就跳过 sparkline 注入
        }
    }

    // ------------------------------------------------------------
    // helpers
    // ------------------------------------------------------------
    private static XSSFRow ensureRow(XSSFSheet sheet, int r) {
        XSSFRow row = sheet.getRow(r);
        return row != null ? row : sheet.createRow(r);
    }

    private static XSSFCell ensureCell(XSSFRow row, int c) {
        XSSFCell cell = row.getCell(c);
        return cell != null ? cell : row.createCell(c);
    }

    private static void writeRow(XSSFSheet sheet, int r, Object... values) {
        writeRowAt(sheet, r, 0, values);
    }

    private static void writeRowAt(XSSFSheet sheet, int r, int startCol, Object... values) {
        Row row = sheet.getRow(r);
        if (row == null) {
            row = sheet.createRow(r);
        }
        for (int i = 0; i < values.length; i++) {
            org.apache.poi.ss.usermodel.Cell c = row.getCell(startCol + i);
            if (c == null) {
                c = row.createCell(startCol + i);
            }
            Object v = values[i];
            if (v instanceof Number) {
                c.setCellValue(((Number) v).doubleValue());
            } else {
                c.setCellValue(String.valueOf(v));
            }
        }
    }

    // Unused to silence javac for SpreadsheetVersion import (kept for potential future fixture work).
    private static void touchImports() {
        SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
        if (v == null) {
            throw new IllegalStateException();
        }
    }
}
