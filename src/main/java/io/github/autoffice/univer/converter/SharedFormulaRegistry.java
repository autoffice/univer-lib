package io.github.autoffice.univer.converter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

/**
 * 共享公式注册器 / Shared-formula registry for Univer's {@code si} feature.
 *
 * <p>读路径：扫描 POI 公式单元格，按表达式分组生成 {@code si}，保持同一组所有单元格共享同一个 {@code si}，
 * 主单元格选取组内右下角的位置。</p>
 *
 * <p>Read path: scan POI formula cells, group by expression and assign one {@code si} per group,
 * picking the bottom-right cell as master.</p>
 *
 * <p>写路径：登记所有 (row, col, si, f) 元组，写入时按 si 分组，确保主单元格位于组内右下角。</p>
 *
 * <p>Write path: register (row, col, si, formula) tuples, then on write group by si and ensure
 * the master cell sits at the bottom-right of the group.</p>
 */
public final class SharedFormulaRegistry {

    /** 读路径：表达式到 si 的映射（按 sheet 索引分桶） / Read path: (sheetIndex, expression) to si. */
    private final Map<String, String> readFormulaToSi = new HashMap<>();

    /** 读路径：si 到主单元格坐标的映射 / Read path: si to master [row, col]. */
    private final Map<String, int[]> readSiToMasterCoord = new HashMap<>();

    /** 写路径：si 到所有共享单元格条目的映射 / Write path: si to all shared cell entries. */
    private final Map<String, List<WriteEntry>> writeSiToEntries = new LinkedHashMap<>();

    /** 写路径条目 / Write-path entry holding (sheetIndex, row, col, formula). */
    static final class WriteEntry {
        final int sheetIndex;
        final int row;
        final int col;
        final String formula;

        WriteEntry(int sheetIndex, int row, int col, String formula) {
            this.sheetIndex = sheetIndex;
            this.row = row;
            this.col = col;
            this.formula = formula;
        }
    }

    /** 组合 sheetIndex 与表达式为读路径查询键 / Build the composite read-path key. */
    private static String readKey(int sheetIndex, String expr) {
        return sheetIndex + "\t" + expr;
    }

    /** 读路径：登记单元格公式并返回其 si / Read path: register a cell formula and return its si. */
    public String registerRead(int sheetIndex, int row, int col, String formula) {
        String expr = stripLeadingEquals(formula);
        String key = readKey(sheetIndex, expr);
        String si = readFormulaToSi.get(key);
        if (si == null) {
            si = UUID.randomUUID().toString();
            readFormulaToSi.put(key, si);
            readSiToMasterCoord.put(si, new int[]{row, col});
        } else {
            int[] current = readSiToMasterCoord.get(si);
            // 右下角优先：行号大者胜，同行则列号大者胜 / Bottom-right wins: larger row, tiebreak by larger col.
            if (row > current[0] || (row == current[0] && col > current[1])) {
                readSiToMasterCoord.put(si, new int[]{row, col});
            }
        }
        return si;
    }

    /** 读路径：根据 si 查找主公式表达式（跨 sheet 唯一）/ Read path: lookup master formula expression by si. */
    public Optional<String> masterFormulaOf(String si) {
        for (Map.Entry<String, String> e : readFormulaToSi.entrySet()) {
            if (e.getValue().equals(si)) {
                String key = e.getKey();
                int tab = key.indexOf('\t');
                return Optional.of(tab >= 0 ? key.substring(tab + 1) : key);
            }
        }
        return Optional.empty();
    }

    /** 读路径：根据 si 查找主单元格坐标 / Read path: lookup master cell coords [row, col] by si. */
    public int[] masterCoordOf(String si) {
        return readSiToMasterCoord.get(si);
    }

    /** 写路径：登记共享同一 si 的单元格 / Write path: register a cell sharing the given si. */
    public void registerWrite(int sheetIndex, int row, int col, String si, String formula) {
        writeSiToEntries.computeIfAbsent(si, k -> new ArrayList<>())
                .add(new WriteEntry(sheetIndex, row, col, formula));
    }

    /** 写路径：登记完成后把公式写入工作簿，确保主单元格位于右下角 / Write path: flush registered cells to the workbook, ensuring master at bottom-right. */
    public void applyOnWorkbook(XSSFWorkbook wb) {
        for (Map.Entry<String, List<WriteEntry>> e : writeSiToEntries.entrySet()) {
            List<WriteEntry> entries = e.getValue();
            if (entries.isEmpty()) {
                continue;
            }
            // 排序后末尾即右下角主单元格 / Sort ascending so the last element is the bottom-right master.
            entries.sort((a, b) -> a.row != b.row
                    ? Integer.compare(a.row, b.row)
                    : Integer.compare(a.col, b.col));
            WriteEntry master = entries.get(entries.size() - 1);
            String formula = stripLeadingEquals(master.formula);
            for (WriteEntry entry : entries) {
                XSSFSheet sheet = wb.getSheetAt(entry.sheetIndex);
                XSSFRow row = sheet.getRow(entry.row);
                if (row == null) {
                    row = sheet.createRow(entry.row);
                }
                XSSFCell cell = row.getCell(entry.col);
                if (cell == null) {
                    cell = row.createCell(entry.col);
                }
                cell.setCellFormula(formula);
            }
        }
    }

    /** 去除公式前导等号 / Strip the leading '=' sign from a formula. */
    private static String stripLeadingEquals(String f) {
        return f != null && f.startsWith("=") ? f.substring(1) : f;
    }
}
