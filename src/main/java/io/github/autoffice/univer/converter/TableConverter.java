package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.TableStyleInfo;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Excel 表格（Table / ListObject）转换器：在 POI {@link XSSFTable} 与 Univer
 * {@code SHEET_TABLE_PLUGIN} 插件资源之间做 best-effort 双向映射。
 *
 * <p>Univer 侧 per-sheet payload 结构：
 * <pre>
 * {
 *   "<tableId>": {
 *     "name": "Table1",
 *     "displayName": "Table1",
 *     "ref": "A1:C10",
 *     "startRow": 0, "endRow": 9, "startColumn": 0, "endColumn": 2,
 *     "headerRowCount": 1,
 *     "totalsRowCount": 0,
 *     "columns": [{ "id": 1, "name": "Col1" }, ...],
 *     "style": {
 *       "name": "TableStyleMedium2",
 *       "showColumnStripes": false,
 *       "showRowStripes": true,
 *       "showFirstColumn": false,
 *       "showLastColumn": false
 *     }
 *   }
 * }
 * </pre>
 *
 * <p>读取时只采集上述字段；写入时通过 {@link XSSFSheet#createTable(AreaReference)}
 * 重新建立表结构，再覆盖列名与样式。自定义 filterColumn、计算列公式、connection 等高级
 * 字段不尝试恢复——这些会通过 sidecar 基线/unknown-extras 机制自然保留。
 *
 * <p>Bridges POI {@link XSSFTable} (Excel Table / ListObject) and Univer's
 * {@code SHEET_TABLE_PLUGIN} resource with best-effort fidelity. Advanced fields
 * (per-column calculated formulas, table connections, complex filter criteria)
 * are not rebuilt on write and rely on sidecar preservation.
 */
public final class TableConverter {

    /** Univer 表格插件名 / Univer table plugin resource name. */
    public static final String PLUGIN_NAME = "SHEET_TABLE_PLUGIN";

    private TableConverter() {
    }

    // ============================================================
    // 读路径 / Read path: POI -> Univer {tableId: payload} JSON
    // ============================================================

    /**
     * 读取 sheet 的所有表格，按 tableId 建立映射；无表格时返回空对象。
     * Read all tables from a sheet into a {tableId: payload} map.
     */
    public static ObjectNode readSheetTables(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (sheet == null) {
            return out;
        }
        List<XSSFTable> tables = sheet.getTables();
        if (tables == null || tables.isEmpty()) {
            return out;
        }
        for (XSSFTable table : tables) {
            if (table == null) {
                continue;
            }
            String key = tableKey(table);
            out.set(key, toTableJson(table, mapper));
        }
        return out;
    }

    private static String tableKey(XSSFTable table) {
        CTTable ct = table.getCTTable();
        if (ct != null && ct.getId() > 0L) {
            return String.valueOf(ct.getId());
        }
        String name = table.getName();
        return name == null || name.isEmpty() ? table.getDisplayName() : name;
    }

    private static ObjectNode toTableJson(XSSFTable table, ObjectMapper mapper) {
        ObjectNode node = mapper.createObjectNode();
        CTTable ct = table.getCTTable();
        if (ct != null && ct.getId() > 0L) {
            node.put("id", ct.getId());
        }
        String name = table.getName();
        if (name != null) {
            node.put("name", name);
        }
        String display = table.getDisplayName();
        if (display != null) {
            node.put("displayName", display);
        }
        String ref = ct != null ? ct.getRef() : null;
        if (ref == null || ref.isEmpty()) {
            AreaReference area = table.getCellReferences();
            if (area != null) {
                ref = area.formatAsString();
            }
        }
        if (ref != null && !ref.isEmpty()) {
            node.put("ref", ref);
        }
        node.put("startRow", table.getStartRowIndex());
        node.put("endRow", table.getEndRowIndex());
        node.put("startColumn", table.getStartColIndex());
        node.put("endColumn", table.getEndColIndex());
        node.put("headerRowCount", table.getHeaderRowCount());
        node.put("totalsRowCount", table.getTotalsRowCount());

        List<XSSFTableColumn> columns = table.getColumns();
        if (columns != null && !columns.isEmpty()) {
            ArrayNode arr = mapper.createArrayNode();
            for (XSSFTableColumn col : columns) {
                if (col == null) {
                    continue;
                }
                ObjectNode c = mapper.createObjectNode();
                c.put("id", col.getId());
                if (col.getName() != null) {
                    c.put("name", col.getName());
                }
                arr.add(c);
            }
            node.set("columns", arr);
        }
        TableStyleInfo style = table.getStyle();
        if (style != null) {
            ObjectNode s = mapper.createObjectNode();
            if (style.getName() != null) {
                s.put("name", style.getName());
            }
            s.put("showColumnStripes", style.isShowColumnStripes());
            s.put("showRowStripes", style.isShowRowStripes());
            s.put("showFirstColumn", style.isShowFirstColumn());
            s.put("showLastColumn", style.isShowLastColumn());
            node.set("style", s);
        }
        return node;
    }

    // ============================================================
    // 写路径 / Write path: Univer JSON -> POI XSSFTable
    // ============================================================

    /**
     * 把 Univer 表格资源写回 sheet。对每个 table 调用 {@link XSSFSheet#createTable}
     * 新建，再覆盖列名和样式。顺序按 payload 字段顺序。
     * Apply a Univer table resource to a sheet by creating a fresh {@link XSSFTable}
     * for each entry, then overlaying column names and style info.
     */
    public static void writeSheetTables(XSSFSheet sheet, JsonNode payload) {
        if (sheet == null || payload == null || !payload.isObject() || payload.size() == 0) {
            return;
        }
        // 避免重复写入：若 sheet 已存在同名/同 ref 的 table，跳过对应 payload。
        // Skip payloads whose target already exists to avoid duplicate tables.
        List<String> existingRefs = new ArrayList<>();
        List<String> existingNames = new ArrayList<>();
        if (sheet.getTables() != null) {
            for (XSSFTable t : sheet.getTables()) {
                if (t == null) {
                    continue;
                }
                if (t.getName() != null) {
                    existingNames.add(t.getName());
                }
                CTTable ct = t.getCTTable();
                if (ct != null && ct.getRef() != null) {
                    existingRefs.add(ct.getRef());
                }
            }
        }

        Iterator<Map.Entry<String, JsonNode>> it = payload.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> e = it.next();
            JsonNode tableNode = e.getValue();
            if (tableNode == null || !tableNode.isObject()) {
                continue;
            }
            writeOneTable(sheet, tableNode, existingRefs, existingNames);
        }
    }

    private static void writeOneTable(XSSFSheet sheet, JsonNode tableNode,
                                      List<String> existingRefs, List<String> existingNames) {
        AreaReference area = resolveArea(tableNode);
        if (area == null) {
            return;
        }
        String refStr = area.formatAsString();
        String name = tableNode.path("name").asText(null);
        if (existingRefs.contains(refStr) || (name != null && existingNames.contains(name))) {
            return;
        }

        XSSFTable table = sheet.createTable(area);
        CTTable ct = table.getCTTable();
        if (ct == null) {
            return;
        }
        ct.setRef(refStr);
        if (name != null && !name.isEmpty()) {
            ct.setName(name);
        }
        String display = tableNode.path("displayName").asText(null);
        if (display != null && !display.isEmpty()) {
            ct.setDisplayName(display);
        } else if (name != null && !name.isEmpty()) {
            ct.setDisplayName(name);
        }
        if (tableNode.hasNonNull("headerRowCount")) {
            int v = tableNode.get("headerRowCount").asInt();
            if (v >= 0) {
                ct.setHeaderRowCount(v);
            }
        }
        if (tableNode.hasNonNull("totalsRowCount")) {
            int v = tableNode.get("totalsRowCount").asInt();
            if (v > 0) {
                ct.setTotalsRowCount(v);
                ct.setTotalsRowShown(true);
            }
        }

        applyColumns(ct, tableNode.path("columns"));
        applyStyle(table, tableNode.path("style"));

        // 注册到 existing lists，避免同一次写入过程中的重复 ref。
        existingRefs.add(refStr);
        if (name != null) {
            existingNames.add(name);
        }
    }

    private static AreaReference resolveArea(JsonNode tableNode) {
        String ref = tableNode.path("ref").asText(null);
        if (ref != null && !ref.isEmpty()) {
            try {
                return new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
            } catch (IllegalArgumentException ignored) {
                // 落回数值字段 / fall through to numeric fields
            }
        }
        if (tableNode.hasNonNull("startRow") && tableNode.hasNonNull("endRow")
                && tableNode.hasNonNull("startColumn") && tableNode.hasNonNull("endColumn")) {
            try {
                CellRangeAddress range = new CellRangeAddress(
                        tableNode.get("startRow").asInt(),
                        tableNode.get("endRow").asInt(),
                        tableNode.get("startColumn").asInt(),
                        tableNode.get("endColumn").asInt());
                return new AreaReference(range.formatAsString(), SpreadsheetVersion.EXCEL2007);
            } catch (RuntimeException ignored) {
                return null;
            }
        }
        return null;
    }

    private static void applyColumns(CTTable ct, JsonNode columnsNode) {
        if (columnsNode == null || !columnsNode.isArray() || columnsNode.size() == 0) {
            return;
        }
        CTTableColumns ctCols = ct.getTableColumns();
        if (ctCols == null) {
            return;
        }
        CTTableColumn[] cols = ctCols.getTableColumnArray();
        int limit = Math.min(cols.length, columnsNode.size());
        for (int i = 0; i < limit; i++) {
            JsonNode c = columnsNode.get(i);
            if (c == null || !c.isObject()) {
                continue;
            }
            if (c.hasNonNull("id")) {
                cols[i].setId(c.get("id").asLong());
            }
            String nm = c.path("name").asText(null);
            if (nm != null && !nm.isEmpty()) {
                cols[i].setName(nm);
            }
        }
    }

    private static void applyStyle(XSSFTable table, JsonNode styleNode) {
        if (styleNode == null || !styleNode.isObject() || styleNode.size() == 0) {
            return;
        }
        CTTable ct = table.getCTTable();
        if (ct == null) {
            return;
        }
        CTTableStyleInfo ctStyle = ct.isSetTableStyleInfo() ? ct.getTableStyleInfo() : ct.addNewTableStyleInfo();
        String name = styleNode.path("name").asText(null);
        if (name != null && !name.isEmpty()) {
            ctStyle.setName(name);
        }
        if (styleNode.hasNonNull("showColumnStripes")) {
            ctStyle.setShowColumnStripes(styleNode.get("showColumnStripes").asBoolean());
        }
        if (styleNode.hasNonNull("showRowStripes")) {
            ctStyle.setShowRowStripes(styleNode.get("showRowStripes").asBoolean());
        }
        if (styleNode.hasNonNull("showFirstColumn")) {
            ctStyle.setShowFirstColumn(styleNode.get("showFirstColumn").asBoolean());
        }
        if (styleNode.hasNonNull("showLastColumn")) {
            ctStyle.setShowLastColumn(styleNode.get("showLastColumn").asBoolean());
        }
        // 如果 sidecar 提供 style 且 XSSFTableStyleInfo 与 ct 同步，刷新 TableStyleInfo 视图。
        // Refresh the XSSFTableStyleInfo wrapper so TableStyleInfo getters stay in sync with ct.
        if (table.getStyle() instanceof XSSFTableStyleInfo) {
            // no-op：XSSFTableStyleInfo 内部持有 ctStyle 引用，setter 已即时生效
        }
    }
}
