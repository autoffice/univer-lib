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

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 性能测试：大数据量场景下的转换性能。
 * Performance test for large data scenarios.
 *
 * <p>默认禁用，需要时手动启用运行。
 * Disabled by default, enable manually when needed.
 */
public class PerformanceTest {

    @Test
    @Disabled("性能测试，需要时手动启用 / Performance test, enable manually when needed")
    public void testLargeWorkbook() throws Exception {
        int rows = 10000;
        int cols = 100;

        System.out.println("=== Performance Test: " + rows + " rows × " + cols + " cols ===");

        // 生成测试数据 / generate test data
        long startGen = System.currentTimeMillis();
        IWorkbookData wb = generateLargeWorkbook(rows, cols);
        long endGen = System.currentTimeMillis();
        System.out.println("Data generation: " + (endGen - startGen) + " ms");

        // 写入 xlsx / write to xlsx
        long startWrite = System.currentTimeMillis();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(wb, out);
        long endWrite = System.currentTimeMillis();
        System.out.println("Write to xlsx: " + (endWrite - startWrite) + " ms");
        System.out.println("Output size: " + (out.size() / 1024) + " KB");

        // 读取 xlsx / read from xlsx
        long startRead = System.currentTimeMillis();
        ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
        IWorkbookData wb2 = UniverXlsx.read(in);
        long endRead = System.currentTimeMillis();
        System.out.println("Read from xlsx: " + (endRead - startRead) + " ms");

        // 验证数据完整性 / verify data integrity
        IWorksheetData sheet = wb2.getSheets().get("Sheet1");
        int cellCount = sheet.getCellData().values().stream()
                .mapToInt(row -> row.size())
                .sum();
        System.out.println("Cells read: " + cellCount);
        System.out.println("Styles count: " + wb2.getStyles().size());

        // 总结 / summary
        long total = (endWrite - startWrite) + (endRead - startRead);
        System.out.println("Total time: " + total + " ms");
    }

    private IWorkbookData generateLargeWorkbook(int rows, int cols) {
        IWorkbookData wb = new IWorkbookData();
        wb.setSheets(new LinkedHashMap<>());
        wb.setStyles(new LinkedHashMap<>());

        IWorksheetData sheet = new IWorksheetData();
        sheet.setId("Sheet1");
        sheet.setName("Sheet1");
        sheet.setCellData(new LinkedHashMap<>());

        // 创建几种不同的样式 / create several different styles
        IStyleData style1 = new IStyleData().setFs(12).setFf("Arial");
        IStyleData style2 = new IStyleData().setFs(14).setBl(io.github.autoffice.univer.model.BooleanNumber.TRUE);
        IStyleData style3 = new IStyleData().setFs(10).setIt(io.github.autoffice.univer.model.BooleanNumber.TRUE);

        // 填充数据 / populate data
        for (int r = 0; r < rows; r++) {
            Map<Integer, ICellData> rowMap = new LinkedHashMap<>();
            for (int c = 0; c < cols; c++) {
                ICellData cell = new ICellData();
                // 交替使用不同类型的数据 / alternate between different data types
                if (c % 3 == 0) {
                    cell.setV("Text_" + r + "_" + c);
                    cell.setT(io.github.autoffice.univer.model.CellValueType.STRING);
                } else if (c % 3 == 1) {
                    cell.setV((double) (r * cols + c));
                    cell.setT(io.github.autoffice.univer.model.CellValueType.NUMBER);
                } else {
                    cell.setV(r % 2 == 0);
                    cell.setT(io.github.autoffice.univer.model.CellValueType.BOOLEAN);
                }
                // 交替使用不同样式 / alternate between different styles
                if (r % 3 == 0) {
                    cell.setS(style1);
                } else if (r % 3 == 1) {
                    cell.setS(style2);
                } else {
                    cell.setS(style3);
                }
                rowMap.put(c, cell);
            }
            sheet.getCellData().put(r, rowMap);
        }

        wb.getSheets().put("Sheet1", sheet);
        return wb;
    }
}
