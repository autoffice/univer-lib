package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import io.github.autoffice.univer.util.JsonMapper;
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

    /** Univer 条件格式插件名 / Univer conditional formatting plugin resource name. */
    private static final String CF_PLUGIN_NAME = "SHEET_CONDITIONAL_FORMATTING_PLUGIN";
    /** Univer 数据验证插件名 / Univer data validation plugin resource name. */
    private static final String DV_PLUGIN_NAME = DataValidationConverter.PLUGIN_NAME;
    /** Univer 数据透视表插件名 / Univer pivot table plugin resource name. */
    private static final String PIVOT_PLUGIN_NAME = PivotTableConverter.PIVOT_PLUGIN_NAME;
    /** Univer 单元格批注插件名 / Univer cell note plugin resource name. */
    private static final String NOTE_PLUGIN_NAME = "SHEET_NOTE_PLUGIN";
    /** Univer 图片/drawing 插件名 / Univer drawing plugin resource name. */
    private static final String DRAWING_PLUGIN_NAME = "SHEET_DRAWING_PLUGIN";
    /** Univer 图表插件名 / Univer chart plugin resource name. */
    private static final String CHART_PLUGIN_NAME = "SHEET_CHART_PLUGIN";
    /** Univer 形状插件名 / Univer shape plugin resource name. */
    private static final String SHAPE_PLUGIN_NAME = "SHEET_SHAPE_PLUGIN";
    /** Univer 迷你图插件名 / Univer sparkline plugin resource name. */
    private static final String SPARKLINE_PLUGIN_NAME = "SHEET_SPARKLINE_PLUGIN";

    private final UniverXlsxOptions opts;

    public WorkbookConverter(UniverXlsxOptions opts) {
        this.opts = opts;
    }

    /** 把 IWorkbookData 转为 POI 工作簿。/ Convert IWorkbookData to POI workbook. */
    public XSSFWorkbook toXlsx(IWorkbookData src) throws io.github.autoffice.univer.UniverXlsxUnsupportedFeatureException {
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

        // 条件格式：按 sheetId 分桶，写入时回填
        Map<String, JsonNode> cfBySheetId = extractResourceBySheetId(src, CF_PLUGIN_NAME);
        // 数据验证：每个 sheet 对应 {ruleId: IDataValidationRule}
        Map<String, JsonNode> dvBySheetId = extractResourceBySheetId(src, DV_PLUGIN_NAME);
        // 透视表：每个 sheet 对应 [pivotPayload...]
        Map<String, JsonNode> pivotBySheetId = extractResourceBySheetId(src, PIVOT_PLUGIN_NAME);
        // 单元格批注：同样按 sheetId 分桶；每个 sheet 对应 {row:{col:note}} 对象
        Map<String, JsonNode> notesBySheetId = extractResourceBySheetId(src, NOTE_PLUGIN_NAME);
        // 图片：每个 sheet 对应 {data:{drawingId:image},order:[drawingId]}
        Map<String, JsonNode> drawingsBySheetId = extractResourceBySheetId(src, DRAWING_PLUGIN_NAME);
        // 图表：每个 sheet 对应 {data:{chartId:item},order:[chartId]}，item.rawXml 为原 OOXML CTChartSpace
        Map<String, JsonNode> chartsBySheetId = extractResourceBySheetId(src, CHART_PLUGIN_NAME);
        // 形状：与 drawing 插件结构一致，但走 simpleShape / connector 通道
        Map<String, JsonNode> shapesBySheetId = extractResourceBySheetId(src, SHAPE_PLUGIN_NAME);
        // 迷你图：保守保留原 worksheet.extLst XML，按 sheetId 分桶
        Map<String, JsonNode> sparklinesBySheetId = extractResourceBySheetId(src, SPARKLINE_PLUGIN_NAME);

        // 建立 sheetId -> 实际写入的 XSSFSheet 映射，便于最终写 CF
        Map<String, XSSFSheet> sheetIdToXssf = new LinkedHashMap<>();
        // 建立 sheetId -> 最终 sheet 名称映射，供超链接写路径翻译 #gid=<id>&range=...
        Map<String, String> sheetIdToName = new LinkedHashMap<>();

        // 第一轮：仅创建 sheet 并登记名称，确保 hyperlink 写入时所有目标 sheet 都已就位。
        // First pass: create sheets and record names so cross-sheet hyperlinks can resolve any target.
        for (String sid : order) {
            IWorksheetData ws = src.getSheets().get(sid);
            if (ws == null) {
                continue;
            }
            String sheetName = ws.getName() != null ? ws.getName() : sid;
            sheetName = wb.getSheet(sheetName) == null ? sheetName : sheetName + "_" + sid;
            XSSFSheet sheet = wb.createSheet(sheetName);
            sheetIdToXssf.put(sid, sheet);
            sheetIdToName.put(sid, sheet.getSheetName());
        }
        wsc.setSheetIdMaps(sheetIdToName, null);

        // 第二轮：实际写入内容与可见性。
        // Second pass: write cell contents and visibility now that hyperlink targets are known.
        for (String sid : order) {
            IWorksheetData ws = src.getSheets().get(sid);
            if (ws == null) {
                continue;
            }
            XSSFSheet sheet = sheetIdToXssf.get(sid);
            if (sheet == null) {
                continue;
            }
            ws = resolveStyles(src, ws);
            wsc.writeSheet(sheet, ws);
            if (ws.getHidden() == BooleanNumber.TRUE) {
                wb.setSheetVisibility(wb.getSheetIndex(sheet), SheetVisibility.HIDDEN);
            }
        }
        sfr.applyOnWorkbook(wb);

        // 回写条件格式
        if (!cfBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : cfBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    ConditionalFormattingConverter.writeSheetCF(sh, e.getValue());
                }
            }
        }
        // 回写数据验证
        if (!dvBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : dvBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    DataValidationConverter.writeSheetDataValidations(sh, e.getValue(), opts);
                }
            }
        }
        // 回写透视表
        if (!pivotBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : pivotBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    PivotTableConverter.writeSheetPivotTables(wb, sh, e.getValue(), sheetIdToXssf);
                }
            }
        }
        // 回写单元格批注
        if (!notesBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : notesBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    CommentConverter.writeSheetComments(sh, e.getValue());
                }
            }
        }
        // 回写图片
        if (!drawingsBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : drawingsBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    PictureConverter.writeSheetPictures(sh, e.getValue());
                }
            }
        }
        // 回写非图片形状（rect、connector、流程图块等）
        if (!shapesBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : shapesBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    AdvancedDrawingConverter.writeSheetShapes(sh, e.getValue());
                }
            }
        }
        // 回写图表（基于 sidecar 中的原始 CTChartSpace XML）
        if (!chartsBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : chartsBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    AdvancedDrawingConverter.writeSheetCharts(sh, e.getValue());
                }
            }
        }
        // 回写迷你图（worksheet.extLst 原 XML 还原）
        if (!sparklinesBySheetId.isEmpty()) {
            for (Map.Entry<String, JsonNode> e : sparklinesBySheetId.entrySet()) {
                XSSFSheet sh = sheetIdToXssf.get(e.getKey());
                if (sh != null) {
                    AdvancedDrawingConverter.writeSheetSparkline(sh, e.getValue());
                }
            }
        }
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
        // 把 xlsx 实际 sheet 名也补进映射，让 'Sheet'!A1 形式的内部 hyperlink 能解析到 sid。
        // Add every workbook sheet name as well, so hyperlink translation can resolve any target.
        Map<String, String> nameToSheetId = new LinkedHashMap<>(nameToId);
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetAt(i).getSheetName();
            nameToSheetId.putIfAbsent(name, nameToId.getOrDefault(name, name));
        }
        wsc.setSheetIdMaps(null, nameToSheetId);

        List<String> order = new ArrayList<>();
        ObjectMapper mapper = JsonMapper.get();
        // Univer CF 插件期望按 sheetId 分组的规则数组
        ObjectNode cfBySheetId = mapper.createObjectNode();
        // Univer Data Validation 插件期望 {sheetId: {ruleId: IDataValidationRule}}
        ObjectNode dvBySheetId = mapper.createObjectNode();
        // Univer Pivot 插件期望 {sheetId: [pivotPayload...]}
        ObjectNode pivotBySheetId = mapper.createObjectNode();
        // Univer Note 插件期望 {sheetId: {row: {col: ISheetNote}}} 嵌套结构
        ObjectNode notesBySheetId = mapper.createObjectNode();
        // Univer Drawing 插件期望 {sheetId: {data:{drawingId:image}, order:[drawingId]}}
        ObjectNode drawingsBySheetId = mapper.createObjectNode();
        // Univer Chart 插件保留原 CTChartSpace XML
        ObjectNode chartsBySheetId = mapper.createObjectNode();
        // Univer Shape 插件保留 simpleShape / connector 元数据
        ObjectNode shapesBySheetId = mapper.createObjectNode();
        // Univer Sparkline 插件保留 worksheet.extLst 原 XML
        ObjectNode sparklinesBySheetId = mapper.createObjectNode();
        String unitId = out.getId();
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

            // 收集条件格式
            ArrayNode cfRules = ConditionalFormattingConverter.readSheetCF(sheet, mapper);
            if (cfRules.size() > 0) {
                cfBySheetId.set(sid, cfRules);
            }
            // 收集数据验证
            ObjectNode validations = DataValidationConverter.readSheetDataValidations(sheet, mapper);
            if (validations.size() > 0) {
                dvBySheetId.set(sid, validations);
            }
            // 收集透视表
            ArrayNode pivotTables = PivotTableConverter.readSheetPivotTables(wb, sheet, mapper,
                    unitId == null ? "" : unitId, sid);
            if (pivotTables.size() > 0) {
                pivotBySheetId.set(sid, pivotTables);
            }
            // 收集单元格批注
            ObjectNode noteMap = CommentConverter.readSheetComments(sheet, mapper);
            if (noteMap.size() > 0) {
                notesBySheetId.set(sid, noteMap);
            }
            // 收集浮动图片
            ObjectNode pictureMap = PictureConverter.readSheetPictures(sheet, mapper,
                    unitId == null ? "" : unitId, sid);
            if (pictureMap.path("data").size() > 0) {
                drawingsBySheetId.set(sid, pictureMap);
            }
            // 收集图表
            ObjectNode chartMap = AdvancedDrawingConverter.readSheetCharts(sheet, mapper,
                    unitId == null ? "" : unitId, sid);
            if (chartMap.path("data").size() > 0) {
                chartsBySheetId.set(sid, chartMap);
            }
            // 收集非图片形状
            ObjectNode shapeMap = AdvancedDrawingConverter.readSheetShapes(sheet, mapper,
                    unitId == null ? "" : unitId, sid);
            if (shapeMap.path("data").size() > 0) {
                shapesBySheetId.set(sid, shapeMap);
            }
            // 收集迷你图：worksheet.extLst 原 XML
            ObjectNode sparklineMap = AdvancedDrawingConverter.readSheetSparkline(sheet, mapper);
            if (sparklineMap.size() > 0) {
                sparklinesBySheetId.set(sid, sparklineMap);
            }
        }
        // 合并 / 追加 SHEET_CONDITIONAL_FORMATTING_PLUGIN 资源
        if (cfBySheetId.size() > 0) {
            mergeResourceBySheetId(out, CF_PLUGIN_NAME, cfBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_DATA_VALIDATION_PLUGIN 资源
        if (dvBySheetId.size() > 0) {
            mergeResourceBySheetId(out, DV_PLUGIN_NAME, dvBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_PIVOT_TABLE_PLUGIN 资源
        if (pivotBySheetId.size() > 0) {
            mergeResourceBySheetId(out, PIVOT_PLUGIN_NAME, pivotBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_NOTE_PLUGIN 资源
        if (notesBySheetId.size() > 0) {
            mergeResourceBySheetId(out, NOTE_PLUGIN_NAME, notesBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_DRAWING_PLUGIN 资源
        if (drawingsBySheetId.size() > 0) {
            mergeResourceBySheetId(out, DRAWING_PLUGIN_NAME, drawingsBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_CHART_PLUGIN 资源
        if (chartsBySheetId.size() > 0) {
            mergeResourceBySheetId(out, CHART_PLUGIN_NAME, chartsBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_SHAPE_PLUGIN 资源
        if (shapesBySheetId.size() > 0) {
            mergeResourceBySheetId(out, SHAPE_PLUGIN_NAME, shapesBySheetId, mapper);
        }
        // 合并 / 追加 SHEET_SPARKLINE_PLUGIN 资源
        if (sparklinesBySheetId.size() > 0) {
            mergeResourceBySheetId(out, SPARKLINE_PLUGIN_NAME, sparklinesBySheetId, mapper);
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

    // ============================================================
    // Plugin resource helpers (SHEET_CONDITIONAL_FORMATTING_PLUGIN / SHEET_NOTE_PLUGIN / ...)
    // ============================================================

    /**
     * 从 {@code src.resources} 中抽取指定插件资源的 per-sheet 数据（{@code {sheetId: ...}}）。
     * <p>resources 在反序列化后常为 {@code List<Map<String,Object>>}，data 字段是 JSON 字符串
     * 或已被二次反序列化成 {@code Map}；两种情况都要兼容。
     * <p>Extract per-sheet data for a plugin resource; works for any {sheetId: payload} schema.
     */
    @SuppressWarnings("unchecked")
    private Map<String, JsonNode> extractResourceBySheetId(IWorkbookData src, String pluginName) {
        Map<String, JsonNode> result = new LinkedHashMap<>();
        Object res = src.getResources();
        if (!(res instanceof List)) {
            return result;
        }
        ObjectMapper mapper = JsonMapper.get();
        for (Object item : (List<Object>) res) {
            if (!(item instanceof Map)) {
                continue;
            }
            Map<String, Object> entry = (Map<String, Object>) item;
            if (!pluginName.equals(String.valueOf(entry.get("name")))) {
                continue;
            }
            Object data = entry.get("data");
            JsonNode bySheet;
            try {
                if (data instanceof String) {
                    String str = ((String) data).trim();
                    if (str.isEmpty()) {
                        continue;
                    }
                    bySheet = mapper.readTree(str);
                } else if (data != null) {
                    bySheet = mapper.valueToTree(data);
                } else {
                    continue;
                }
            } catch (Exception e) {
                continue;
            }
            if (bySheet == null || !bySheet.isObject()) {
                continue;
            }
            bySheet.fields().forEachRemaining(f -> {
                if (f.getValue() != null && !f.getValue().isNull()) {
                    result.put(f.getKey(), f.getValue());
                }
            });
        }
        return result;
    }

    /**
     * 合并 / 追加 plugin 资源项；若已存在同名资源，把新值按 sheetId 合并进原 data，否则新增一条。
     * Merge or append a plugin resource entry keyed by {@code pluginName}.
     */
    @SuppressWarnings("unchecked")
    private void mergeResourceBySheetId(IWorkbookData out, String pluginName,
                                        ObjectNode bySheetId, ObjectMapper mapper) {
        List<Object> resources;
        Object existingRes = out.getResources();
        if (existingRes instanceof List) {
            resources = (List<Object>) existingRes;
        } else {
            resources = new ArrayList<>();
        }
        // 查找已有条目
        Map<String, Object> target = null;
        for (Object item : resources) {
            if (item instanceof Map) {
                Map<String, Object> m = (Map<String, Object>) item;
                if (pluginName.equals(String.valueOf(m.get("name")))) {
                    target = m;
                    break;
                }
            }
        }
        ObjectNode merged = mapper.createObjectNode();
        if (target != null) {
            Object data = target.get("data");
            try {
                JsonNode existing = null;
                if (data instanceof String && !((String) data).trim().isEmpty()) {
                    existing = mapper.readTree((String) data);
                } else if (data != null) {
                    existing = mapper.valueToTree(data);
                }
                if (existing != null && existing.isObject()) {
                    existing.fields().forEachRemaining(f -> merged.set(f.getKey(), f.getValue()));
                }
            } catch (Exception ignored) {
                // 坏数据直接忽略
            }
        }
        bySheetId.fields().forEachRemaining(f -> merged.set(f.getKey(), f.getValue()));

        try {
            String dataStr = mapper.writeValueAsString(merged);
            if (target != null) {
                target.put("data", dataStr);
            } else {
                Map<String, Object> entry = new LinkedHashMap<>();
                entry.put("name", pluginName);
                entry.put("data", dataStr);
                resources.add(entry);
            }
        } catch (Exception ignored) {
            return;
        }
        out.setResources(resources);
    }
}
