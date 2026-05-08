package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作簿级双向转换器。
 * Workbook-level converter between IWorkbookData and XSSFWorkbook.
 */
public final class WorkbookConverter {
    private final UniverXlsxOptions opts;

    public WorkbookConverter(UniverXlsxOptions opts) {
        this.opts = opts;
    }

    /** 把 IWorkbookData 转为 POI 工作簿。/ Convert IWorkbookData to POI workbook. */
    public XSSFWorkbook toXlsx(IWorkbookData src) {
        XSSFWorkbook wb = new XSSFWorkbook();
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, opts);

        if (src == null) {
            return wb;
        }

        List<String> order = src.getSheetOrder() != null && !src.getSheetOrder().isEmpty()
                ? src.getSheetOrder()
                : new ArrayList<>(src.getSheets().keySet());

        for (String sid : order) {
            IWorksheetData ws = src.getSheets().get(sid);
            if (ws == null) {
                continue;
            }
            ws = resolveStyles(src, ws);
            String sheetName = ws.getName() != null ? ws.getName() : sid;
            // 确保工作表名唯一 / ensure unique sheet name
            sheetName = wb.getSheet(sheetName) == null ? sheetName : sheetName + "_" + sid;
            XSSFSheet sheet = wb.createSheet(sheetName);
            wsc.writeSheet(sheet, ws);
            if (ws.getHidden() == BooleanNumber.TRUE) {
                wb.setSheetVisibility(wb.getSheetIndex(sheet), SheetVisibility.HIDDEN);
            }
        }
        sfr.applyOnWorkbook(wb);
        return wb;
    }

    /** 从 POI 工作簿读取（可结合边车基线）。/ Read from POI workbook, optionally merging with sidecar baseline. */
    public IWorkbookData fromXlsx(XSSFWorkbook wb, IWorkbookData sidecarBaseline) {
        IWorkbookData out = sidecarBaseline != null ? sidecarBaseline : new IWorkbookData();
        if (out.getSheets() == null) {
            out.setSheets(new LinkedHashMap<>());
        }
        if (out.getStyles() == null) {
            out.setStyles(new LinkedHashMap<>());
        }

        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, opts);

        // 构建 sidecar 工作表名到 id 的反向索引，解决 xlsx 按名、sidecar 按 id 存储的错配
        // Build name → id map from sidecar baseline so xlsx-by-name lookups hit sidecar entries.
        Map<String, String> nameToId = new LinkedHashMap<>();
        for (Map.Entry<String, IWorksheetData> e : out.getSheets().entrySet()) {
            IWorksheetData v = e.getValue();
            if (v != null && v.getName() != null) {
                nameToId.put(v.getName(), e.getKey());
            }
        }

        List<String> order = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            String sid = nameToId.getOrDefault(sheetName, sheetName);
            IWorksheetData ws = wsc.readSheet(sheet);
            ws.setId(sid);
            if (wb.getSheetVisibility(i) == SheetVisibility.HIDDEN) {
                ws.setHidden(BooleanNumber.TRUE);
            }
            // 合并：xlsx 内容字段优先，sidecar 提供辅助字段
            IWorksheetData existing = out.getSheets().get(sid);
            if (existing != null) {
                mergeSheetData(existing, ws);
                out.getSheets().put(sid, existing);
            } else {
                out.getSheets().put(sid, ws);
            }
            order.add(sid);
        }
        if (out.getSheetOrder() == null || out.getSheetOrder().isEmpty()) {
            out.setSheetOrder(order);
        }
        if (out.getLocale() == null) {
            out.setLocale(opts.getLocale());
        }
        // 回填 styles：cell.s 用字符串 id 引用，需补齐映射表以便消费方解析
        // Populate styles: cell.s references style id strings, so ensure the map has those keys.
        Map<String, IStyleData> observed = sc.getStyleRegistry();
        if (!observed.isEmpty()) {
            for (Map.Entry<String, IStyleData> e : observed.entrySet()) {
                out.getStyles().putIfAbsent(e.getKey(), e.getValue());
            }
        }
        return out;
    }

    /** 解引用 styleId → IStyleData，便于 cell 写入。/ Resolve string style ids in cellData to inline IStyleData. */
    private IWorksheetData resolveStyles(IWorkbookData wb, IWorksheetData ws) {
        if (wb.getStyles() == null || wb.getStyles().isEmpty()) {
            return ws;
        }
        if (ws.getCellData() == null) {
            return ws;
        }
        for (Map.Entry<Integer, Map<Integer, ICellData>> rowE : ws.getCellData().entrySet()) {
            for (Map.Entry<Integer, ICellData> colE : rowE.getValue().entrySet()) {
                ICellData cd = colE.getValue();
                if (cd != null && cd.getS() instanceof String) {
                    IStyleData inline = wb.getStyles().get((String) cd.getS());
                    if (inline != null) {
                        cd.setS(inline);
                    }
                }
            }
        }
        return ws;
    }

    /** 合并：xlsx 读出的 ws 覆盖 sidecar 中可由 xlsx 推断的字段。/ Merge: xlsx-derived overrides sidecar for content fields. */
    private void mergeSheetData(IWorksheetData baseline, IWorksheetData fromXlsx) {
        if (fromXlsx.getCellData() != null && !fromXlsx.getCellData().isEmpty()) {
            mergeCellData(baseline, fromXlsx);
        }
        if (fromXlsx.getMergeData() != null && !fromXlsx.getMergeData().isEmpty()) {
            baseline.setMergeData(fromXlsx.getMergeData());
        }
        if (fromXlsx.getRowData() != null && !fromXlsx.getRowData().isEmpty()) {
            baseline.setRowData(fromXlsx.getRowData());
        }
        if (fromXlsx.getColumnData() != null && !fromXlsx.getColumnData().isEmpty()) {
            baseline.setColumnData(fromXlsx.getColumnData());
        }
        if (baseline.getName() == null) {
            baseline.setName(fromXlsx.getName());
        }
        if (baseline.getHidden() == null) {
            baseline.setHidden(fromXlsx.getHidden());
        }
        if (baseline.getFreeze() == null) {
            baseline.setFreeze(fromXlsx.getFreeze());
        }
        if (baseline.getShowGridlines() == null) {
            baseline.setShowGridlines(fromXlsx.getShowGridlines());
        }
        if (baseline.getRightToLeft() == null) {
            baseline.setRightToLeft(fromXlsx.getRightToLeft());
        }
        if (baseline.getTabColor() == null) {
            baseline.setTabColor(fromXlsx.getTabColor());
        }
        if (baseline.getDefaultColumnWidth() == null) {
            baseline.setDefaultColumnWidth(fromXlsx.getDefaultColumnWidth());
        }
        if (baseline.getDefaultRowHeight() == null) {
            baseline.setDefaultRowHeight(fromXlsx.getDefaultRowHeight());
        }
    }

    /**
     * 逐单元格合并：xlsx 覆盖值/类型/公式/样式，保留 sidecar 独有字段（富文本 p、自定义 custom）。
     * Cell-level merge: xlsx overrides value/type/formula/style, but sidecar-only
     * fields (rich text {@code p}, {@code custom}) are preserved when xlsx has no replacement.
     */
    private void mergeCellData(IWorksheetData baseline, IWorksheetData fromXlsx) {
        if (baseline.getCellData() == null) {
            baseline.setCellData(new LinkedHashMap<>());
        }
        Map<Integer, Map<Integer, ICellData>> base = baseline.getCellData();
        for (Map.Entry<Integer, Map<Integer, ICellData>> rowE : fromXlsx.getCellData().entrySet()) {
            Integer rowIdx = rowE.getKey();
            Map<Integer, ICellData> baseRow = base.computeIfAbsent(rowIdx, k -> new LinkedHashMap<>());
            for (Map.Entry<Integer, ICellData> colE : rowE.getValue().entrySet()) {
                Integer colIdx = colE.getKey();
                ICellData xc = colE.getValue();
                if (xc == null) {
                    continue;
                }
                ICellData bc = baseRow.get(colIdx);
                if (bc == null) {
                    baseRow.put(colIdx, xc);
                    continue;
                }
                // 覆盖 xlsx 可推断字段 / override fields derivable from xlsx
                bc.setV(xc.getV());
                bc.setT(xc.getT());
                bc.setF(xc.getF());
                bc.setSi(xc.getSi());
                if (xc.getS() != null) {
                    bc.setS(xc.getS());
                }
                // 保留 baseline.p（富文本）与 baseline.custom：xlsx 无法完整表达，sidecar 值继续生效
                // Preserve baseline.p and baseline.custom: xlsx can't fully express; sidecar wins.
            }
        }
    }
}
