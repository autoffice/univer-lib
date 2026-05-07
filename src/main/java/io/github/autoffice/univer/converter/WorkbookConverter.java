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

        List<String> order = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            IWorksheetData ws = wsc.readSheet(sheet);
            String sid = ws.getId() != null ? ws.getId() : sheet.getSheetName();
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
            baseline.setCellData(fromXlsx.getCellData());
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
}
