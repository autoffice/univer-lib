package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IColumnData;
import io.github.autoffice.univer.model.IFreeze;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IRowData;
import io.github.autoffice.univer.model.IWorksheetData;
import io.github.autoffice.univer.util.ColorUtils;
import io.github.autoffice.univer.util.LengthUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作表转换器：在 IWorksheetData 与 POI XSSFSheet 之间双向映射。
 * Worksheet converter between IWorksheetData and POI XSSFSheet.
 * <p>
 * 组合 CellConverter / RichTextConverter / SharedFormulaRegistry / StyleConverter 完成单元格级工作。
 * Orchestrates CellConverter / RichTextConverter / SharedFormulaRegistry / StyleConverter.
 */
public final class WorksheetConverter {

    private final XSSFWorkbook wb;
    private final StyleConverter styles;
    private final CellConverter cells;
    private final RichTextConverter rich;
    private final SharedFormulaRegistry formulas;
    @SuppressWarnings("unused")
    private final UniverXlsxOptions opts;

    public WorksheetConverter(XSSFWorkbook wb,
                              StyleConverter styles,
                              CellConverter cells,
                              RichTextConverter rich,
                              SharedFormulaRegistry formulas,
                              UniverXlsxOptions opts) {
        this.wb = wb;
        this.styles = styles;
        this.cells = cells;
        this.rich = rich;
        this.formulas = formulas;
        this.opts = opts;
    }

    // ============================================================
    // 写路径 / Write path
    // ============================================================

    /**
     * 把 IWorksheetData 写入 POI sheet。
     * Write IWorksheetData into POI sheet.
     */
    public void writeSheet(XSSFSheet sheet, IWorksheetData src) {
        if (src == null) {
            return;
        }
        writeSheetAttrs(sheet, src);
        writeCellData(sheet, src);
        writeRowData(sheet, src);
        writeColumnData(sheet, src);
        writeMergedRegions(sheet, src);
        writeFreeze(sheet, src);
    }

    /** 写工作表级属性 / write sheet-level attributes. */
    private void writeSheetAttrs(XSSFSheet sheet, IWorksheetData src) {
        // tabColor
        if (src.getTabColor() != null) {
            byte[] argb = ColorUtils.rgbHexToArgb(src.getTabColor());
            if (argb != null) {
                sheet.setTabColor(new XSSFColor(argb, null));
            }
        }
        // defaultColumnWidth (px)
        if (src.getDefaultColumnWidth() != null) {
            sheet.setDefaultColumnWidth((int) LengthUtils.pxToChars(src.getDefaultColumnWidth()));
        }
        // defaultRowHeight (px)
        if (src.getDefaultRowHeight() != null) {
            sheet.setDefaultRowHeightInPoints((float) LengthUtils.pxToPoints(src.getDefaultRowHeight()));
        }
        // showGridlines
        if (src.getShowGridlines() == BooleanNumber.FALSE) {
            sheet.setDisplayGridlines(false);
        }
        // rightToLeft
        if (src.getRightToLeft() == BooleanNumber.TRUE) {
            sheet.setRightToLeft(true);
        }
        // zoomRatio
        if (src.getZoomRatio() != null) {
            int zoom = (int) (src.getZoomRatio() * 100);
            if (zoom > 0) {
                sheet.setZoom(zoom);
            }
        }
    }

    /** 写单元格数据 / write cell data matrix. */
    private void writeCellData(XSSFSheet sheet, IWorksheetData src) {
        if (src.getCellData() == null || src.getCellData().isEmpty()) {
            return;
        }
        int sheetIndex = wb.getSheetIndex(sheet);
        for (Map.Entry<Integer, Map<Integer, ICellData>> rowEntry : src.getCellData().entrySet()) {
            int rowIdx = rowEntry.getKey();
            Map<Integer, ICellData> rowMap = rowEntry.getValue();
            if (rowMap == null || rowMap.isEmpty()) {
                continue;
            }
            XSSFRow poiRow = sheet.getRow(rowIdx);
            if (poiRow == null) {
                poiRow = sheet.createRow(rowIdx);
            }
            for (Map.Entry<Integer, ICellData> colEntry : rowMap.entrySet()) {
                int colIdx = colEntry.getKey();
                ICellData cellData = colEntry.getValue();
                if (cellData == null) {
                    continue;
                }
                XSSFCell cell = poiRow.getCell(colIdx);
                if (cell == null) {
                    cell = poiRow.createCell(colIdx);
                }

                if (cellData.getP() != null) {
                    // 富文本 / rich text
                    XSSFRichTextString rts = rich.toPoi(cellData.getP());
                    cell.setCellValue(rts);
                    applyStyleOnly(cell, cellData);
                } else if (cellData.getF() != null && cellData.getSi() != null) {
                    // 共享公式 / shared formula — 登记到 registry，由其 apply 时统一写入
                    formulas.registerWrite(sheetIndex, rowIdx, colIdx, cellData.getSi(), cellData.getF());
                    applyStyleOnly(cell, cellData);
                } else {
                    // 普通单元格 / plain cell
                    cells.writeCell(cell, cellData);
                }
            }
        }
    }

    /** 仅应用样式（富文本 / si 公式用） / apply style only for rich-text or si-formula cells. */
    private void applyStyleOnly(XSSFCell cell, ICellData src) {
        Object styleObj = src.getS();
        if (styleObj instanceof io.github.autoffice.univer.model.IStyleData) {
            cell.setCellStyle(styles.toPoiStyle((io.github.autoffice.univer.model.IStyleData) styleObj));
        }
        // 字符串 id 由 WorkbookConverter 层负责解析 / string style id resolved by higher layer
    }

    /** 写行数据 / write row data. */
    private void writeRowData(XSSFSheet sheet, IWorksheetData src) {
        if (src.getRowData() == null || src.getRowData().isEmpty()) {
            return;
        }
        for (Map.Entry<Integer, IRowData> e : src.getRowData().entrySet()) {
            int rowIdx = e.getKey();
            IRowData rd = e.getValue();
            if (rd == null) {
                continue;
            }
            XSSFRow poiRow = sheet.getRow(rowIdx);
            if (poiRow == null) {
                poiRow = sheet.createRow(rowIdx);
            }
            if (rd.getH() != null) {
                poiRow.setHeightInPoints((float) LengthUtils.pxToPoints(rd.getH()));
            }
            if (rd.getHd() == BooleanNumber.TRUE) {
                poiRow.setZeroHeight(true);
            }
        }
    }

    /** 写列数据 / write column data. */
    private void writeColumnData(XSSFSheet sheet, IWorksheetData src) {
        if (src.getColumnData() == null || src.getColumnData().isEmpty()) {
            return;
        }
        for (Map.Entry<Integer, IColumnData> e : src.getColumnData().entrySet()) {
            int colIdx = e.getKey();
            IColumnData cd = e.getValue();
            if (cd == null) {
                continue;
            }
            if (cd.getW() != null) {
                sheet.setColumnWidth(colIdx, (int) (LengthUtils.pxToChars(cd.getW()) * 256));
            }
            if (cd.getHd() == BooleanNumber.TRUE) {
                sheet.setColumnHidden(colIdx, true);
            }
        }
    }

    /** 写合并区域 / write merged regions. */
    private void writeMergedRegions(XSSFSheet sheet, IWorksheetData src) {
        List<IRange> merges = src.getMergeData();
        if (merges == null || merges.isEmpty()) {
            return;
        }
        for (IRange r : merges) {
            if (r == null || r.getStartRow() == null || r.getEndRow() == null
                    || r.getStartColumn() == null || r.getEndColumn() == null) {
                continue;
            }
            sheet.addMergedRegion(new CellRangeAddress(
                    r.getStartRow(), r.getEndRow(), r.getStartColumn(), r.getEndColumn()));
        }
    }

    /** 写冻结窗格 / write freeze pane. */
    private void writeFreeze(XSSFSheet sheet, IWorksheetData src) {
        IFreeze f = src.getFreeze();
        if (f == null) {
            return;
        }
        int startRow = f.getStartRow() == null ? -1 : f.getStartRow();
        int startColumn = f.getStartColumn() == null ? -1 : f.getStartColumn();
        int xSplit = f.getXSplit() == null ? 0 : f.getXSplit();
        int ySplit = f.getYSplit() == null ? 0 : f.getYSplit();
        // 全为 "无冻结" 语义时跳过 / skip when everything denotes no-freeze
        if (startRow == -1 && startColumn == -1 && xSplit == 0 && ySplit == 0) {
            return;
        }
        sheet.createFreezePane(xSplit, ySplit, Math.max(0, startColumn), Math.max(0, startRow));
    }

    // ============================================================
    // 读路径 / Read path
    // ============================================================

    /**
     * 从 POI sheet 读取 IWorksheetData。
     * Read IWorksheetData from POI sheet.
     */
    public IWorksheetData readSheet(XSSFSheet sheet) {
        IWorksheetData dst = new IWorksheetData()
                .setId(sheet.getSheetName())
                .setName(sheet.getSheetName());

        readSheetAttrs(sheet, dst);
        readCellsAndRows(sheet, dst);
        readColumnData(sheet, dst);
        readMergedRegions(sheet, dst);
        readFreeze(sheet, dst);
        // rowCount / columnCount 是 Univer 渲染网格必需字段；缺失会导致前端显示空白表
        // rowCount / columnCount are required for Univer to render the grid; missing them shows blank
        computeDimensions(sheet, dst);

        return dst;
    }

    /** 推断 rowCount / columnCount：取实际使用范围再留一些空行空列。 */
    private void computeDimensions(XSSFSheet sheet, IWorksheetData dst) {
        int maxRow = sheet.getLastRowNum();
        int maxCol = 0;
        for (org.apache.poi.ss.usermodel.Row row : sheet) {
            short last = row.getLastCellNum();
            if (last > maxCol) {
                maxCol = last;
            }
        }
        // 合并区域也可能扩展边界 / merged regions can extend bounds
        if (dst.getMergeData() != null) {
            for (IRange r : dst.getMergeData()) {
                if (r.getEndRow() != null && r.getEndRow() > maxRow) {
                    maxRow = r.getEndRow();
                }
                if (r.getEndColumn() != null && r.getEndColumn() >= maxCol) {
                    maxCol = r.getEndColumn() + 1;
                }
            }
        }
        // 给用户留一些空行空列用于编辑；最小值参考 Univer 默认（1000 行 × 26 列）
        // leave headroom for editing; minimums match Univer defaults (1000 rows × 26 cols)
        int rowCount = Math.max(maxRow + 20, 100);
        int columnCount = Math.max(maxCol + 5, 26);
        dst.setRowCount(rowCount);
        dst.setColumnCount(columnCount);
    }

    /** 读工作表级属性 / read sheet-level attributes. */
    private void readSheetAttrs(XSSFSheet sheet, IWorksheetData dst) {
        XSSFColor tabColor = sheet.getTabColor();
        if (tabColor != null) {
            byte[] argb = tabColor.getARGB();
            if (argb != null && argb.length >= 4) {
                String hex = ColorUtils.argbToRgbHex(argb);
                if (hex != null) {
                    dst.setTabColor(hex);
                }
            }
        }
        dst.setDefaultColumnWidth(LengthUtils.charsToPx((double) sheet.getDefaultColumnWidth()));
        dst.setDefaultRowHeight(LengthUtils.pointsToPx(sheet.getDefaultRowHeightInPoints()));
        dst.setShowGridlines(sheet.isDisplayGridlines() ? BooleanNumber.TRUE : BooleanNumber.FALSE);
        dst.setRightToLeft(sheet.isRightToLeft() ? BooleanNumber.TRUE : BooleanNumber.FALSE);
        // zoomRatio: POI 无公开 getZoom，故跳过 / skip zoom read (no public getter)
    }

    /** 读单元格数据和行数据（合并为单次遍历）/ read cell data and row data in single traversal. */
    private void readCellsAndRows(XSSFSheet sheet, IWorksheetData dst) {
        int sheetIndex = wb.getSheetIndex(sheet);
        float defaultH = sheet.getDefaultRowHeightInPoints();

        for (org.apache.poi.ss.usermodel.Row row : sheet) {
            if (!(row instanceof XSSFRow)) {
                continue;
            }
            XSSFRow xRow = (XSSFRow) row;
            int rowIdx = xRow.getRowNum();

            // 收集行元数据（高度、隐藏）/ collect row metadata (height, hidden)
            float h = xRow.getHeightInPoints();
            boolean hidden = xRow.getZeroHeight();
            boolean custom = Math.abs(h - defaultH) > 0.01f;
            if (custom || hidden) {
                IRowData rd = new IRowData();
                if (custom) {
                    rd.setH(LengthUtils.pointsToPx(h));
                }
                if (hidden) {
                    rd.setHd(BooleanNumber.TRUE);
                }
                dst.getRowData().put(rowIdx, rd);
            }

            // 遍历单元格 / traverse cells
            for (org.apache.poi.ss.usermodel.Cell cell : xRow) {
                if (!(cell instanceof XSSFCell)) {
                    continue;
                }
                XSSFCell xCell = (XSSFCell) cell;
                int colIdx = xCell.getColumnIndex();
                ICellData data = cells.readCell(xCell);
                if (data == null) {
                    continue;
                }
                // 富文本检测：STRING 类型且有多个 formatting runs 时转为 IDocumentData
                // Rich text detection: if STRING cell has multiple formatting runs, convert to IDocumentData
                if (xCell.getCellType() == CellType.STRING) {
                    XSSFRichTextString rts = xCell.getRichStringCellValue();
                    if (rts != null && rts.numFormattingRuns() > 1) {
                        data.setP(rich.fromPoi(rts));
                    }
                }
                // 公式：追加 si / formula: attach si
                if (xCell.getCellType() == CellType.FORMULA && data.getF() != null) {
                    String si = formulas.registerRead(sheetIndex, rowIdx, colIdx, data.getF());
                    data.setSi(si);
                }
                dst.getCellData()
                        .computeIfAbsent(rowIdx, k -> new LinkedHashMap<>())
                        .put(colIdx, data);
            }
        }
    }

    /** 读列数据 / read column data (only columns with custom width or hidden). */
    private void readColumnData(XSSFSheet sheet, IWorksheetData dst) {
        // 直接从 CTCols 读取列元数据，避免遍历所有行 / read column metadata directly from CTCols
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols[] colsArr =
                sheet.getCTWorksheet().getColsArray();
        if (colsArr == null || colsArr.length == 0) {
            return;
        }
        int defaultWidth256 = sheet.getDefaultColumnWidth() * 256;
        for (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols cols : colsArr) {
            for (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol col : cols.getColArray()) {
                // CTCol 的 min/max 定义了列范围（1-based，需要转为 0-based）
                // CTCol min/max define column range (1-based, convert to 0-based)
                int minCol = (int) col.getMin() - 1;
                int maxCol = (int) col.getMax() - 1;
                for (int i = minCol; i <= maxCol; i++) {
                    int w = sheet.getColumnWidth(i);
                    boolean hidden = sheet.isColumnHidden(i);
                    boolean customW = w != defaultWidth256;
                    if (!customW && !hidden) {
                        continue;
                    }
                    IColumnData cd = new IColumnData();
                    if (customW) {
                        cd.setW(LengthUtils.charsToPx(w / 256.0));
                    }
                    if (hidden) {
                        cd.setHd(BooleanNumber.TRUE);
                    }
                    dst.getColumnData().put(i, cd);
                }
            }
        }
    }

    /** 读合并区域 / read merged regions. */
    private void readMergedRegions(XSSFSheet sheet, IWorksheetData dst) {
        int n = sheet.getNumMergedRegions();
        if (n <= 0) {
            return;
        }
        java.util.List<IRange> list = new java.util.ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            list.add(new IRange()
                    .setStartRow(r.getFirstRow())
                    .setEndRow(r.getLastRow())
                    .setStartColumn(r.getFirstColumn())
                    .setEndColumn(r.getLastColumn()));
        }
        dst.setMergeData(list);
    }

    /** 读冻结窗格 / read freeze pane. */
    private void readFreeze(XSSFSheet sheet, IWorksheetData dst) {
        PaneInformation info = sheet.getPaneInformation();
        if (info == null || !info.isFreezePane()) {
            return;
        }
        IFreeze f = new IFreeze()
                .setXSplit((int) info.getVerticalSplitPosition())
                .setYSplit((int) info.getHorizontalSplitPosition())
                .setStartRow((int) info.getHorizontalSplitTopRow())
                .setStartColumn((int) info.getVerticalSplitLeftColumn());
        dst.setFreeze(f);
    }
}
