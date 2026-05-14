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
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IColumnData;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IFreeze;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IRowData;
import io.github.autoffice.univer.model.IWorksheetData;
import io.github.autoffice.univer.util.ColorUtils;
import io.github.autoffice.univer.util.LengthUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 工作表转换器：在 IWorksheetData 与 POI XSSFSheet 之间双向映射。
 * Worksheet converter between IWorksheetData and POI XSSFSheet.
 * <p>
 * 组合 CellConverter / RichTextConverter / SharedFormulaRegistry / StyleConverter 完成单元格级工作。
 * Orchestrates CellConverter / RichTextConverter / SharedFormulaRegistry / StyleConverter.
 */
public final class WorksheetConverter {

    private static final Logger LOG = Logger.getLogger(WorksheetConverter.class.getName());

    /** Univer 内部超链接格式：#gid=<sheetId>&range=<a1> / Univer internal hyperlink url shape. */
    private static final Pattern GID_PATTERN = Pattern.compile("^#gid=([^&]+)(?:&range=(.+))?$");
    /** xlsx 原生内部超链接格式：[quoted]sheet!range / xlsx native internal hyperlink shape. */
    private static final Pattern SHEET_REF_PATTERN = Pattern.compile("^'?(.+?)'?!(.+)$");

    private final XSSFWorkbook wb;
    private final StyleConverter styles;
    private final CellConverter cells;
    private final RichTextConverter rich;
    private final SharedFormulaRegistry formulas;
    @SuppressWarnings("unused")
    private final UniverXlsxOptions opts;

    /** sheetId → sheetName（写路径使用，解析 #gid=...）。 */
    private Map<String, String> sheetIdToName = Collections.emptyMap();
    /** sheetName → sheetId（读路径使用，生成 #gid=...）。 */
    private Map<String, String> sheetNameToId = Collections.emptyMap();

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

    /**
     * 注入跨 sheet 超链接解析所需的映射。
     * Provide cross-sheet maps used to translate Univer {@code #gid=} URLs to xlsx
     * {@code Sheet!A1} style references, and vice versa.
     */
    public void setSheetIdMaps(Map<String, String> sheetIdToName, Map<String, String> sheetNameToId) {
        this.sheetIdToName = sheetIdToName == null ? Collections.emptyMap() : sheetIdToName;
        this.sheetNameToId = sheetNameToId == null ? Collections.emptyMap() : sheetNameToId;
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
                    // 特殊情况：p 只是为了承载 cell-level hyperlink 的占位（dataStream="\r\n"，无文本），
                    // 不应强行把 cell 类型从 BLANK 升级成 STRING。仅应用样式与 hyperlink。
                    if (isHyperlinkOnlyDoc(cellData.getP())) {
                        applyStyleOnly(cell, cellData);
                        applyHyperlinkFromDoc(cell, cellData.getP());
                    } else {
                        XSSFRichTextString rts = rich.toPoi(cellData.getP());
                        cell.setCellValue(rts);
                        applyStyleOnly(cell, cellData);
                        applyHyperlinkFromDoc(cell, cellData.getP());
                    }
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

    /**
     * 判断 IDocumentData 是否仅为 hyperlink 占位（没有真实文本）。
     * 读路径在 BLANK cell 上遇到 cell-level hyperlink 时，会构造一个 dataStream 仅为段落终止符 "\r\n" 的
     * IDocumentData 来承载 customRanges。写回时应保持 BLANK，不应强行写 STRING。
     */
    private static boolean isHyperlinkOnlyDoc(IDocumentData doc) {
        if (doc == null || doc.getBody() == null) {
            return false;
        }
        String s = doc.getBody().getDataStream();
        return s == null || s.isEmpty() || "\r\n".equals(s);
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
                // 单元格级 hyperlink → IDocumentData.customRanges
                // Cell-level hyperlink → IDocumentData.customRanges
                XSSFHyperlink link = xCell.getHyperlink();
                if (link != null) {
                    applyHyperlinkToData(data, xCell, link);
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

    // ============================================================
    // 超链接 / Hyperlinks
    // ============================================================

    /**
     * 将 IDocumentData.body.customRanges 中的第一个超链接应用到 xlsx 单元格。
     * xlsx 一个单元格只能承载一个 hyperlink；多链接富文本靠 sidecar 的 p 继续保真。
     * Apply the first customRange hyperlink (if any) to the xlsx cell. xlsx only allows a single
     * cell-level hyperlink; multi-link rich text is still preserved losslessly through the sidecar.
     */
    private void applyHyperlinkFromDoc(XSSFCell cell, IDocumentData doc) {
        String url = RichTextConverter.firstHyperlinkUrl(doc);
        if (url == null) {
            return;
        }
        HyperlinkType type = detectHyperlinkType(url);
        String address = translateUrlForXlsx(url, type);
        if (address == null || address.isEmpty()) {
            return;
        }
        try {
            CreationHelper helper = wb.getCreationHelper();
            Hyperlink h = helper.createHyperlink(type);
            h.setAddress(address);
            cell.setHyperlink(h);
        } catch (IllegalArgumentException e) {
            // Address 解析失败：降级为 sidecar-only 保留；不要中断整体写入
            LOG.fine("skip invalid hyperlink address: " + address + " (" + e.getMessage() + ")");
        }
    }

    /** 把 xlsx cell.hyperlink 合并回 ICellData（生成或补齐 body.customRanges）。 */
    private void applyHyperlinkToData(ICellData data, XSSFCell xCell, XSSFHyperlink link) {
        String url = translateUrlFromXlsx(link, xCell.getSheet());
        if (url == null) {
            return;
        }
        IDocumentData doc = data.getP();
        if (doc == null) {
            doc = rich.fromPoi(xCell.getRichStringCellValue());
            data.setP(doc);
        }
        // 如已通过 sidecar 携带同 url 的 customRange，则不再重复追加
        String existing = RichTextConverter.firstHyperlinkUrl(doc);
        if (url.equals(existing)) {
            return;
        }
        String plain = plainTextOf(doc);
        int end = plain == null ? 0 : plain.length();
        RichTextConverter.attachHyperlink(doc, url, 0, end);
    }

    private static String plainTextOf(IDocumentData doc) {
        if (doc == null || doc.getBody() == null || doc.getBody().getDataStream() == null) {
            return "";
        }
        String s = doc.getBody().getDataStream();
        return s.endsWith("\r\n") ? s.substring(0, s.length() - 2) : s;
    }

    private static HyperlinkType detectHyperlinkType(String url) {
        if (url == null) {
            return HyperlinkType.NONE;
        }
        if (url.startsWith("#")) {
            return HyperlinkType.DOCUMENT;
        }
        String lower = url.toLowerCase();
        if (lower.startsWith("mailto:") || (url.contains("@") && !lower.startsWith("http"))) {
            return HyperlinkType.EMAIL;
        }
        if (lower.startsWith("file:") || lower.startsWith("file://")) {
            return HyperlinkType.FILE;
        }
        return HyperlinkType.URL;
    }

    /** Univer URL → xlsx hyperlink address。内部 #gid= 会被翻译成 'Sheet'!A1 形式。 */
    private String translateUrlForXlsx(String url, HyperlinkType type) {
        if (type != HyperlinkType.DOCUMENT) {
            return url;
        }
        if (url.startsWith("#gid=")) {
            Matcher m = GID_PATTERN.matcher(url);
            if (m.matches()) {
                String sheetId = m.group(1);
                String range = m.group(2);
                String name = sheetIdToName.get(sheetId);
                if (name != null) {
                    String safe = name.replace("'", "''");
                    String target = range == null || range.isEmpty() ? "A1" : range;
                    return "'" + safe + "'!" + target;
                }
            }
            // 未匹配到 sheet 映射，保留原值；部分消费方（本库自己）可直接识别 #gid=。
            return url.substring(1);
        }
        // 已是 xlsx 原生格式（#Sheet!A1 或 #rangeid=...），去掉前导 '#'
        return url.startsWith("#") ? url.substring(1) : url;
    }

    /** xlsx hyperlink → Univer URL。内部 Sheet!A1 会被翻译回 #gid=...&range=...。 */
    private String translateUrlFromXlsx(XSSFHyperlink link, XSSFSheet currentSheet) {
        String address = link.getAddress();
        if (address == null || address.isEmpty()) {
            address = link.getLocation();
        }
        if (address == null || address.isEmpty()) {
            return null;
        }
        HyperlinkType t = link.getType();
        if (t != HyperlinkType.DOCUMENT) {
            return address;
        }
        Matcher m = SHEET_REF_PATTERN.matcher(address);
        if (m.matches()) {
            String rawName = m.group(1).replace("''", "'");
            String range = m.group(2);
            String sid = sheetNameToId.get(rawName);
            if (sid == null && currentSheet != null && rawName.equals(currentSheet.getSheetName())) {
                sid = sheetNameToId.getOrDefault(rawName, rawName);
            }
            if (sid != null) {
                return "#gid=" + sid + (range == null || range.isEmpty() ? "" : "&range=" + range);
            }
        }
        // 未知内部格式，保留原串前加 '#' 以符合 Univer 约定
        return "#" + address;
    }
}
