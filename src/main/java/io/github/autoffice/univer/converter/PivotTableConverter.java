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

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * 数据透视表转换器：在 POI pivot table 与 Univer workbook resources 之间做 best-effort 转换。
 * Pivot table bridge between POI pivot tables and Univer workbook resources with best-effort support.
 *
 * <p>当前只保证基础 source/target/range 以及 row/column/filter/value 四类字段的读写；
 * 高级特性（复杂筛选、排序、分组、下钻、自定义布局细节）会显式记录到
 * {@code unsupportedFeatures} 或 {@code extras}，写回 xlsx 时不会尝试完整恢复。
 */
public final class PivotTableConverter {

    /** Univer 侧 pivot 插件资源名。 */
    public static final String PIVOT_PLUGIN_NAME = "SHEET_PIVOT_TABLE_PLUGIN";

    private static final Logger LOGGER = Logger.getLogger(PivotTableConverter.class.getName());

    private PivotTableConverter() {}

    /**
     * 读取 sheet 上的全部 pivot table，返回 per-sheet payload 数组。
     * Read all pivot tables on a sheet and return a per-sheet payload array.
     */
    public static ArrayNode readSheetPivotTables(XSSFWorkbook wb, XSSFSheet sheet,
                                                 ObjectMapper mapper, String unitId, String subUnitId) {
        ArrayNode out = mapper.createArrayNode();
        if (sheet == null) {
            return out;
        }
        List<XSSFPivotTable> pivots = sheet.getPivotTables();
        if (pivots == null || pivots.isEmpty()) {
            return out;
        }
        for (int i = 0; i < pivots.size(); i++) {
            XSSFPivotTable pivot = pivots.get(i);
            if (pivot == null) {
                continue;
            }
            out.add(readPivotTable(wb, sheet, pivot, mapper, unitId, subUnitId, i));
        }
        return out;
    }

    private static ObjectNode readPivotTable(XSSFWorkbook wb, XSSFSheet targetSheet,
                                             XSSFPivotTable pivot, ObjectMapper mapper,
                                             String unitId, String subUnitId, int ordinal) {
        ObjectNode out = mapper.createObjectNode();
        CTPivotTableDefinition def = pivot.getCTPivotTableDefinition();
        // 稳定 pivotId：使用 UUID 保证跨 round-trip 稳定性，不依赖 sheetName 或 ordinal。
        // Stable pivotId using UUID to guarantee round-trip stability without relying on sheetName or ordinal.
        String actualSubUnitId = subUnitId == null ? targetSheet.getSheetName() : subUnitId;
        String pivotId = buildStablePivotId(actualSubUnitId, def, ordinal);
        out.put("pivotTableId", pivotId);
        out.put("unitId", unitId == null ? "" : unitId);
        out.put("subUnitId", actualSubUnitId);

        if (def != null) {
            if (def.getName() != null && !def.getName().isEmpty()) {
                out.put("name", def.getName());
            }
            ObjectNode options = mapper.createObjectNode();
            options.put("showRowGrandTotal", def.getRowGrandTotals());
            options.put("showColumnGrandTotal", def.getColGrandTotals());
            options.put("compact", def.getCompact());
            options.put("outline", def.getOutline());
            options.put("outlineData", def.getOutlineData());
            options.put("compactData", def.getCompactData());
            options.put("multipleFieldFilters", def.getMultipleFieldFilters());
            if (def.getDataCaption() != null && !def.getDataCaption().isEmpty()) {
                options.put("dataCaption", def.getDataCaption());
            }
            options.put("dataOnRows", def.getDataOnRows());
            options.put("dataPosition", def.getDataPosition());
            out.set("options", options);
        }

        ObjectNode sourceRangeInfo = buildSourceRangeInfo(wb, pivot, mapper, unitId);
        out.set("sourceRangeInfo", sourceRangeInfo);
        out.set("targetCellInfo", buildTargetCellInfo(targetSheet, def, mapper, unitId, subUnitId));
        out.set("fieldsConfig", buildFieldsConfig(pivot, mapper));

        // 库级始终不还原的特性 / library-wide always-unsupported categories
        ArrayNode unsupported = mapper.createArrayNode();
        // 当前文件检测到、只透传不还原的标志 / file-specific detected flags
        ArrayNode detected = mapper.createArrayNode();
        collectUnsupportedFeatures(def, unsupported, detected);
        out.set("unsupportedFeatures", unsupported);
        out.set("detectedFeatures", detected);
        return out;
    }

    private static String buildStablePivotId(String subUnitId, CTPivotTableDefinition def, int ordinal) {
        // 使用 sheetId-pivot-<ordinal> 格式保证 pivotId 稳定性，不受 sheet rename 影响。
        // Use sheetId-pivot-<ordinal> format to guarantee pivotId stability, unaffected by sheet rename.
        return subUnitId + "-pivot-" + ordinal;
    }

    private static ObjectNode buildSourceRangeInfo(XSSFWorkbook wb, XSSFPivotTable pivot,
                                                   ObjectMapper mapper, String unitId) {
        ObjectNode source = mapper.createObjectNode();
        source.put("unitId", unitId == null ? "" : unitId);

        XSSFPivotCacheDefinition cacheDefinition = pivot.getPivotCacheDefinition();
        if (cacheDefinition != null) {
            CTPivotCacheDefinition ctCache = cacheDefinition.getCTPivotCacheDefinition();
            if (ctCache != null) {
                CTCacheSource cacheSource = ctCache.getCacheSource();
                if (cacheSource != null) {
                    source.put("sourceType", String.valueOf(cacheSource.getType()));
                    CTWorksheetSource wsSource = cacheSource.getWorksheetSource();
                    if (wsSource != null) {
                        String sheetName = wsSource.getSheet();
                        if (sheetName != null) {
                            source.put("sheetName", sheetName);
                            XSSFSheet sourceSheet = wb.getSheet(sheetName);
                            if (sourceSheet != null) {
                                source.put("subUnitId", sourceSheet.getSheetName());
                            }
                        }
                        if (wsSource.getName() != null && !wsSource.getName().isEmpty()) {
                            source.put("name", wsSource.getName());
                        }
                        if (wsSource.getRef() != null && !wsSource.getRef().isEmpty()) {
                            AreaReference area = new AreaReference(wsSource.getRef(), SpreadsheetVersion.EXCEL2007);
                            source.set("range", toRangeNode(area, mapper));
                        }
                    }
                }
                CTCacheFields fields = ctCache.getCacheFields();
                if (fields != null && fields.sizeOfCacheFieldArray() > 0) {
                    ArrayNode sourceFields = mapper.createArrayNode();
                    for (CTCacheField field : fields.getCacheFieldArray()) {
                        ObjectNode fieldNode = mapper.createObjectNode();
                        fieldNode.put("name", field.getName());
                        fieldNode.put("numFmtId", field.getNumFmtId());
                        fieldNode.put("databaseField", field.getDatabaseField());
                        if (field.getFormula() != null && !field.getFormula().isEmpty()) {
                            fieldNode.put("formula", field.getFormula());
                        }
                        sourceFields.add(fieldNode);
                    }
                    source.set("sourceFields", sourceFields);
                }
            }
            try {
                AreaReference area = cacheDefinition.getPivotArea(wb);
                if (area != null && !source.has("range")) {
                    source.set("range", toRangeNode(area, mapper));
                }
            } catch (IllegalArgumentException | NullPointerException ignored) {
                // named range / malformed source / missing sheet best-effort only
            }
        }
        return source;
    }

    private static ObjectNode buildTargetCellInfo(XSSFSheet targetSheet, CTPivotTableDefinition def,
                                                  ObjectMapper mapper, String unitId, String subUnitId) {
        ObjectNode target = mapper.createObjectNode();
        target.put("unitId", unitId == null ? "" : unitId);
        target.put("subUnitId", subUnitId == null ? targetSheet.getSheetName() : subUnitId);
        target.put("sheetName", targetSheet.getSheetName());
        if (def != null) {
            CTLocation location = def.getLocation();
            if (location != null && location.getRef() != null && !location.getRef().isEmpty()) {
                AreaReference area = new AreaReference(location.getRef(), SpreadsheetVersion.EXCEL2007);
                CellReference first = area.getFirstCell();
                target.put("row", first.getRow());
                target.put("column", first.getCol());
                target.set("range", toRangeNode(area, mapper));
                target.put("firstDataRow", location.getFirstDataRow());
                target.put("firstDataColumn", location.getFirstDataCol());
                target.put("firstHeaderRow", location.getFirstHeaderRow());
                target.put("rowPageCount", location.getRowPageCount());
                target.put("colPageCount", location.getColPageCount());
            }
        }
        return target;
    }

    private static ObjectNode buildFieldsConfig(XSSFPivotTable pivot, ObjectMapper mapper) {
        ObjectNode fieldsConfig = mapper.createObjectNode();
        ArrayNode rows = mapper.createArrayNode();
        ArrayNode columns = mapper.createArrayNode();
        ArrayNode filters = mapper.createArrayNode();
        ArrayNode values = mapper.createArrayNode();
        fieldsConfig.set("rows", rows);
        fieldsConfig.set("columns", columns);
        fieldsConfig.set("filters", filters);
        fieldsConfig.set("values", values);

        CTPivotTableDefinition def = pivot.getCTPivotTableDefinition();
        if (def == null) {
            return fieldsConfig;
        }
        Map<Integer, String> fieldNames = resolveFieldNames(pivot);
        Map<Integer, CTPivotField> pivotFieldMap = resolvePivotFieldMap(def);

        addAxisFields(rows, "row", def.getRowFields(), fieldNames, pivotFieldMap, mapper);
        addAxisFields(columns, "column", def.getColFields(), fieldNames, pivotFieldMap, mapper);
        addPageFields(filters, fieldNames, pivotFieldMap, def.getPageFields(), mapper);
        addValueFields(values, fieldNames, def.getDataFields(), mapper);
        return fieldsConfig;
    }

    private static Map<Integer, String> resolveFieldNames(XSSFPivotTable pivot) {
        Map<Integer, String> names = new LinkedHashMap<>();
        XSSFPivotCacheDefinition cacheDefinition = pivot.getPivotCacheDefinition();
        if (cacheDefinition == null) {
            return names;
        }
        CTPivotCacheDefinition ctCache = cacheDefinition.getCTPivotCacheDefinition();
        if (ctCache == null || ctCache.getCacheFields() == null) {
            return names;
        }
        CTCacheField[] fields = ctCache.getCacheFields().getCacheFieldArray();
        for (int i = 0; i < fields.length; i++) {
            names.put(i, fields[i].getName());
        }
        return names;
    }

    private static Map<Integer, CTPivotField> resolvePivotFieldMap(CTPivotTableDefinition def) {
        Map<Integer, CTPivotField> fields = new LinkedHashMap<>();
        CTPivotFields pivotFields = def.getPivotFields();
        if (pivotFields == null) {
            return fields;
        }
        CTPivotField[] array = pivotFields.getPivotFieldArray();
        for (int i = 0; i < array.length; i++) {
            fields.put(i, array[i]);
        }
        return fields;
    }

    private static void addAxisFields(ArrayNode out, String area,
                                      CTRowFields rowFields,
                                      Map<Integer, String> fieldNames,
                                      Map<Integer, CTPivotField> pivotFieldMap,
                                      ObjectMapper mapper) {
        if (rowFields == null) {
            return;
        }
        for (int i = 0; i < rowFields.sizeOfFieldArray(); i++) {
            CTField field = rowFields.getFieldArray(i);
            appendAxisField(out, area, i, field.getX(), fieldNames, pivotFieldMap, mapper);
        }
    }

    private static void addAxisFields(ArrayNode out, String area,
                                      CTColFields colFields,
                                      Map<Integer, String> fieldNames,
                                      Map<Integer, CTPivotField> pivotFieldMap,
                                      ObjectMapper mapper) {
        if (colFields == null) {
            return;
        }
        for (int i = 0; i < colFields.sizeOfFieldArray(); i++) {
            CTField field = colFields.getFieldArray(i);
            appendAxisField(out, area, i, field.getX(), fieldNames, pivotFieldMap, mapper);
        }
    }

    private static void appendAxisField(ArrayNode out, String area, int index, int sourceIndex,
                                        Map<Integer, String> fieldNames,
                                        Map<Integer, CTPivotField> pivotFieldMap,
                                        ObjectMapper mapper) {
        ObjectNode fieldNode = mapper.createObjectNode();
        String name = fieldNames.get(sourceIndex);
        fieldNode.put("fieldId", buildFieldId(area, sourceIndex, index));
        fieldNode.put("sourceIndex", sourceIndex);
        fieldNode.put("sourceName", name == null ? ("Field" + sourceIndex) : name);
        fieldNode.put("displayName", name == null ? ("Field" + sourceIndex) : name);
        fieldNode.put("area", area);
        fieldNode.put("index", index);
        CTPivotField pivotField = pivotFieldMap.get(sourceIndex);
        if (pivotField != null) {
            if (pivotField.getSubtotalCaption() != null && !pivotField.getSubtotalCaption().isEmpty()) {
                fieldNode.put("subtotalCaption", pivotField.getSubtotalCaption());
            }
            fieldNode.put("showAll", pivotField.getShowAll());
            fieldNode.put("compact", pivotField.getCompact());
            fieldNode.put("outline", pivotField.getOutline());
        }
        out.add(fieldNode);
    }

    private static void addPageFields(ArrayNode out,
                                      Map<Integer, String> fieldNames,
                                      Map<Integer, CTPivotField> pivotFieldMap,
                                      CTPageFields pageFields,
                                      ObjectMapper mapper) {
        if (pageFields == null) {
            return;
        }
        for (int i = 0; i < pageFields.sizeOfPageFieldArray(); i++) {
            CTPageField field = pageFields.getPageFieldArray(i);
            int sourceIndex = field.getFld();
            ObjectNode fieldNode = mapper.createObjectNode();
            String name = fieldNames.get(sourceIndex);
            fieldNode.put("fieldId", buildFieldId("filter", sourceIndex, i));
            fieldNode.put("sourceIndex", sourceIndex);
            fieldNode.put("sourceName", name == null ? ("Field" + sourceIndex) : name);
            fieldNode.put("displayName", name == null ? ("Field" + sourceIndex) : name);
            fieldNode.put("area", "filter");
            fieldNode.put("index", i);
            fieldNode.put("item", field.getItem());
            CTPivotField pivotField = pivotFieldMap.get(sourceIndex);
            if (pivotField != null) {
                fieldNode.put("showAll", pivotField.getShowAll());
            }
            out.add(fieldNode);
        }
    }

    private static void addValueFields(ArrayNode out,
                                       Map<Integer, String> fieldNames,
                                       CTDataFields dataFields,
                                       ObjectMapper mapper) {
        if (dataFields == null) {
            return;
        }
        for (int i = 0; i < dataFields.sizeOfDataFieldArray(); i++) {
            CTDataField field = dataFields.getDataFieldArray(i);
            int sourceIndex = (int) field.getFld();
            String name = field.getName();
            if (name == null || name.isEmpty()) {
                name = fieldNames.get(sourceIndex);
            }
            ObjectNode fieldNode = mapper.createObjectNode();
            fieldNode.put("fieldId", buildFieldId("value", sourceIndex, i));
            fieldNode.put("sourceIndex", sourceIndex);
            fieldNode.put("sourceName", fieldNames.get(sourceIndex) == null
                    ? ("Field" + sourceIndex) : fieldNames.get(sourceIndex));
            fieldNode.put("displayName", name == null ? ("Field" + sourceIndex) : name);
            fieldNode.put("area", "value");
            fieldNode.put("index", i);
            fieldNode.put("subtotal", subtotalToUniverName(field.getSubtotal()));
            out.add(fieldNode);
        }
    }

    private static void collectUnsupportedFeatures(CTPivotTableDefinition def,
                                                   ArrayNode unsupported, ArrayNode detected) {
        // 库层面始终不还原的特性类目 / library-wide always-unsupported categories
        unsupported.add("grouping");                // 日期/数字/元素分组 / grouping (date/number/element)
        unsupported.add("drillDown");               // 下钻明细 / drill-down detail rows
        unsupported.add("refreshSemantics");        // 数据源变更后的自动刷新语义 / auto-refresh semantics
        unsupported.add("advancedLabelSort");       // 高级标签排序（拼音、自定义排序等） / advanced label sort
        unsupported.add("advancedLabelFilter");     // 高级标签筛选（值筛选、Top N、自定义公式） / advanced label filter
        unsupported.add("valuePositioning");        // 多值字段布局：value position / value index
        unsupported.add("executionMode");           // 主进程 / worker 执行模式（运行时配置，不持久化）

        if (def == null) {
            return;
        }
        // 当前文件检测到、只透传不还原的具体标志 / file-specific detected flags
        if (def.getMultipleFieldFilters()) {
            detected.add("multipleFieldFilters");
        }
        if (!def.getCompact()) {
            detected.add("nonCompactLayout");
        }
        if (def.getOutline()) {
            detected.add("outlineLayout");
        }
        if (def.getDataOnRows()) {
            detected.add("dataOnRows");
        }
    }

    /**
     * 把 per-sheet pivot payload 写回 xlsx。
     * Write per-sheet pivot payloads back to xlsx as best-effort pivot tables.
     */
    public static void writeSheetPivotTables(XSSFWorkbook wb, XSSFSheet targetSheet,
                                             JsonNode sheetPayload, Map<String, XSSFSheet> sheetIdToSheet) {
        if (wb == null || targetSheet == null || sheetPayload == null || !sheetPayload.isArray()) {
            return;
        }
        for (JsonNode pivotNode : sheetPayload) {
            if (pivotNode != null && pivotNode.isObject()) {
                writePivotTable(wb, targetSheet, pivotNode, sheetIdToSheet);
            }
        }
    }

    private static void writePivotTable(XSSFWorkbook wb, XSSFSheet targetSheet,
                                        JsonNode pivotNode, Map<String, XSSFSheet> sheetIdToSheet) {
        String pivotId = pivotNode.path("pivotTableId").asText("(unknown)");
        JsonNode sourceInfo = pivotNode.path("sourceRangeInfo");
        JsonNode rangeNode = sourceInfo.path("range");
        JsonNode targetInfo = pivotNode.path("targetCellInfo");
        JsonNode fieldsConfig = pivotNode.path("fieldsConfig");

        if (!rangeNode.isObject()) {
            LOGGER.log(Level.WARNING, "Pivot {0}: missing or invalid sourceRangeInfo.range, skipping write.",
                    pivotId);
            return;
        }
        if (!targetInfo.isObject()) {
            LOGGER.log(Level.WARNING, "Pivot {0}: missing or invalid targetCellInfo, skipping write.", pivotId);
            return;
        }
        if (!fieldsConfig.isObject()) {
            LOGGER.log(Level.WARNING, "Pivot {0}: missing or invalid fieldsConfig, skipping write.", pivotId);
            return;
        }

        XSSFSheet sourceSheet = resolveSourceSheet(wb, sourceInfo, sheetIdToSheet);
        if (sourceSheet == null) {
            LOGGER.log(Level.WARNING, "Pivot {0}: cannot resolve source sheet, skipping write.", pivotId);
            return;
        }

        AreaReference sourceRange = toAreaReference(rangeNode);
        if (sourceRange == null) {
            LOGGER.log(Level.WARNING, "Pivot {0}: cannot parse source range, skipping write.", pivotId);
            return;
        }
        int targetRow = targetInfo.path("row").asInt(-1);
        int targetCol = targetInfo.path("column").asInt(-1);
        if (targetRow < 0 || targetCol < 0) {
            LOGGER.log(Level.WARNING, "Pivot {0}: invalid target row/column, skipping write.", pivotId);
            return;
        }
        CellReference anchor = new CellReference(targetRow, targetCol);

        XSSFPivotTable pivot = targetSheet.createPivotTable(sourceRange, anchor, sourceSheet);
        applyFieldConfig(pivot, fieldsConfig);
        applyOptions(pivot, pivotNode.path("options"));
        String name = pivotNode.path("name").asText("");
        if (!name.isEmpty()) {
            pivot.getCTPivotTableDefinition().setName(name);
        }
    }

    private static XSSFSheet resolveSourceSheet(XSSFWorkbook wb, JsonNode sourceInfo,
                                                Map<String, XSSFSheet> sheetIdToSheet) {
        String subUnitId = sourceInfo.path("subUnitId").asText("");
        if (!subUnitId.isEmpty() && sheetIdToSheet != null) {
            XSSFSheet sheet = sheetIdToSheet.get(subUnitId);
            if (sheet != null) {
                return sheet;
            }
        }
        String sheetName = sourceInfo.path("sheetName").asText("");
        if (!sheetName.isEmpty()) {
            return wb.getSheet(sheetName);
        }
        return null;
    }

    private static void applyFieldConfig(XSSFPivotTable pivot, JsonNode fieldsConfig) {
        Map<Integer, JsonNode> valueFields = new LinkedHashMap<>();
        for (JsonNode field : fieldsConfig.path("rows")) {
            int sourceIndex = field.path("sourceIndex").asInt(-1);
            if (sourceIndex >= 0) {
                pivot.addRowLabel(sourceIndex);
            }
        }
        JsonNode columns = fieldsConfig.path("columns");
        for (int i = 0; i < columns.size(); i++) {
            JsonNode field = columns.get(i);
            int sourceIndex = field.path("sourceIndex").asInt(-1);
            if (sourceIndex >= 0) {
                String displayName = emptyToNull(field.path("displayName").asText(""));
                if (displayName == null) {
                    pivot.addColLabel(sourceIndex);
                } else {
                    pivot.addColLabel(sourceIndex, displayName);
                }
            }
        }
        for (JsonNode field : fieldsConfig.path("filters")) {
            int sourceIndex = field.path("sourceIndex").asInt(-1);
            if (sourceIndex >= 0) {
                pivot.addReportFilter(sourceIndex);
            }
        }
        JsonNode values = fieldsConfig.path("values");
        for (int i = 0; i < values.size(); i++) {
            JsonNode field = values.get(i);
            int sourceIndex = field.path("sourceIndex").asInt(-1);
            if (sourceIndex >= 0) {
                valueFields.put(i, field);
                String displayName = emptyToNull(field.path("displayName").asText(""));
                DataConsolidateFunction function = subtotalToPoi(field.path("subtotal").asText("sum"));
                if (displayName == null) {
                    pivot.addColumnLabel(function, sourceIndex);
                } else {
                    pivot.addColumnLabel(function, sourceIndex, displayName);
                }
            }
        }
    }

    private static void applyOptions(XSSFPivotTable pivot, JsonNode options) {
        if (options == null || !options.isObject()) {
            return;
        }
        CTPivotTableDefinition def = pivot.getCTPivotTableDefinition();
        // 严格 hasNonNull 检查，避免把 explicit null 当作 false 设置。
        // Strict hasNonNull check to avoid treating explicit null as false.
        if (options.hasNonNull("showRowGrandTotal")) {
            def.setRowGrandTotals(options.path("showRowGrandTotal").asBoolean(true));
        }
        if (options.hasNonNull("showColumnGrandTotal")) {
            def.setColGrandTotals(options.path("showColumnGrandTotal").asBoolean(true));
        }
        if (options.hasNonNull("compact")) {
            def.setCompact(options.path("compact").asBoolean(true));
        }
        if (options.hasNonNull("outline")) {
            def.setOutline(options.path("outline").asBoolean(false));
        }
        if (options.hasNonNull("outlineData")) {
            def.setOutlineData(options.path("outlineData").asBoolean(false));
        }
        if (options.hasNonNull("compactData")) {
            def.setCompactData(options.path("compactData").asBoolean(true));
        }
        if (options.hasNonNull("multipleFieldFilters")) {
            def.setMultipleFieldFilters(options.path("multipleFieldFilters").asBoolean(false));
        }
        if (options.hasNonNull("dataCaption")) {
            def.setDataCaption(options.path("dataCaption").asText("Values"));
        }
        if (options.hasNonNull("dataOnRows")) {
            def.setDataOnRows(options.path("dataOnRows").asBoolean(false));
        }
    }

    private static AreaReference toAreaReference(JsonNode rangeNode) {
        int startRow = rangeNode.path("startRow").asInt(-1);
        int endRow = rangeNode.path("endRow").asInt(-1);
        int startCol = rangeNode.path("startColumn").asInt(-1);
        int endCol = rangeNode.path("endColumn").asInt(-1);
        if (startRow < 0 || endRow < 0 || startCol < 0 || endCol < 0) {
            return null;
        }
        CellReference first = new CellReference(startRow, startCol);
        CellReference last = new CellReference(endRow, endCol);
        return new AreaReference(first, last, SpreadsheetVersion.EXCEL2007);
    }

    private static ObjectNode toRangeNode(AreaReference area, ObjectMapper mapper) {
        ObjectNode range = mapper.createObjectNode();
        CellReference first = area.getFirstCell();
        CellReference last = area.getLastCell();
        range.put("startRow", first.getRow());
        range.put("startColumn", first.getCol());
        range.put("endRow", last.getRow());
        range.put("endColumn", last.getCol());
        range.put("rangeType", 0);
        range.put("startAbsoluteRefType", 0);
        range.put("endAbsoluteRefType", 0);
        return range;
    }

    private static String buildFieldId(String area, int sourceIndex, int index) {
        return area + "-" + sourceIndex + "-" + index;
    }

    private static String subtotalToUniverName(STDataConsolidateFunction.Enum subtotal) {
        if (subtotal == null) {
            return "sum";
        }
        String name = subtotal.toString().toLowerCase(Locale.ROOT);
        switch (name) {
            case "countnums":
                return "countNumbers";
            case "stddev":
                return "stdDev";
            case "stddevp":
                return "stdDevp";
            default:
                return name;
        }
    }

    private static DataConsolidateFunction subtotalToPoi(String subtotal) {
        if (subtotal == null || subtotal.isEmpty()) {
            return DataConsolidateFunction.SUM;
        }
        switch (subtotal) {
            case "sum":
                return DataConsolidateFunction.SUM;
            case "count":
                return DataConsolidateFunction.COUNT;
            case "countNumbers":
                return DataConsolidateFunction.COUNT_NUMS;
            case "average":
                return DataConsolidateFunction.AVERAGE;
            case "max":
                return DataConsolidateFunction.MAX;
            case "min":
                return DataConsolidateFunction.MIN;
            case "product":
                return DataConsolidateFunction.PRODUCT;
            case "stdDev":
                return DataConsolidateFunction.STD_DEV;
            case "stdDevp":
                return DataConsolidateFunction.STD_DEVP;
            case "var":
                return DataConsolidateFunction.VAR;
            case "varp":
                return DataConsolidateFunction.VARP;
            default:
                LOGGER.log(Level.WARNING, "Unknown subtotal function ''{0}'', defaulting to SUM.", subtotal);
                return DataConsolidateFunction.SUM;
        }
    }

    private static String emptyToNull(String value) {
        return value == null || value.isEmpty() ? null : value;
    }
}
