package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * 定义名称转换器：在 POI Workbook 的全局/工作表作用域命名区 与 Univer
 * {@code SHEET_DEFINED_NAME_PLUGIN} 插件数据（{@code {[id]: IDefinedNamesServiceParam}}）之间双向转换。
 *
 * <p>Univer schema 为工作簿级 flat map，每条记录形如：
 * <pre>
 * {
 *   "id": "uuid",
 *   "name": "MyName",
 *   "formulaOrRefString": "Sheet1!$A$1",
 *   "comment": "...",        // optional
 *   "localSheetId": "sid",   // optional，缺省表示 workbook scope
 *   "hidden": false           // optional
 * }
 * </pre>
 */
public final class DefinedNameConverter {

    /** Univer 定义名称插件名 / Univer defined name plugin resource name. */
    public static final String PLUGIN_NAME = "SHEET_DEFINED_NAME_PLUGIN";

    private DefinedNameConverter() {}

    // ============================================================
    // 读路径 / Read path: POI -> Univer {id -> definedName}
    // ============================================================

    /**
     * 读取 workbook 的全部命名区，输出 Univer 定义名称 map。
     * Read all named ranges from the workbook as a Univer defined-name map.
     */
    public static ObjectNode readWorkbookDefinedNames(XSSFWorkbook wb, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (wb == null) {
            return out;
        }
        List<XSSFName> names = wb.getAllNames();
        if (names == null || names.isEmpty()) {
            return out;
        }
        for (XSSFName name : names) {
            if (name == null) {
                continue;
            }
            // 过滤 Excel 内置名称（以 "_xlnm." 开头），由 POI 自己管理，不属 Univer 语义。
            String raw = name.getNameName();
            if (raw == null || raw.isEmpty() || raw.startsWith("_xlnm.")) {
                continue;
            }
            String id = UUID.randomUUID().toString();
            ObjectNode item = mapper.createObjectNode();
            item.put("id", id);
            item.put("name", raw);
            String formula = name.getRefersToFormula();
            item.put("formulaOrRefString", formula == null ? "" : formula);
            if (name.getComment() != null && !name.getComment().isEmpty()) {
                item.put("comment", name.getComment());
            }
            int sheetIdx = name.getSheetIndex();
            if (sheetIdx >= 0) {
                String sheetName = wb.getSheetName(sheetIdx);
                if (sheetName != null) {
                    item.put("localSheetId", sheetName);
                }
            }
            if (name.isHidden()) {
                item.put("hidden", true);
            }
            out.set(id, item);
        }
        return out;
    }

    // ============================================================
    // 写路径 / Write path: Univer map -> POI named ranges
    // ============================================================

    /**
     * 把 Univer 定义名称 map 写回 POI workbook。
     * Apply a Univer defined-name map back to the POI workbook as named ranges.
     */
    public static void writeWorkbookDefinedNames(XSSFWorkbook wb, JsonNode namesNode) {
        if (wb == null || namesNode == null || !namesNode.isObject() || namesNode.size() == 0) {
            return;
        }
        Iterator<Map.Entry<String, JsonNode>> it = namesNode.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> entry = it.next();
            JsonNode item = entry.getValue();
            if (item == null || !item.isObject()) {
                continue;
            }
            String nameStr = item.path("name").asText("");
            String formula = item.path("formulaOrRefString").asText("");
            if (nameStr.isEmpty() || formula.isEmpty()) {
                continue;
            }
            // 同名冲突时先删除旧的，避免 POI 抛异常（xlsx 不允许同作用域同名）
            Name existing = wb.getName(nameStr);
            if (existing != null) {
                try {
                    wb.removeName(existing);
                } catch (Exception ignored) {
                    // 删除失败则跳过，避免破坏整体写入
                    continue;
                }
            }
            try {
                Name n = wb.createName();
                n.setNameName(nameStr);
                n.setRefersToFormula(formula);
                String localSheet = item.path("localSheetId").asText("");
                if (!localSheet.isEmpty()) {
                    int idx = wb.getSheetIndex(localSheet);
                    if (idx >= 0) {
                        n.setSheetIndex(idx);
                    }
                }
                String comment = item.path("comment").asText("");
                if (!comment.isEmpty()) {
                    n.setComment(comment);
                }
                // 注意：POI 5.2.5 的 XSSFName 不暴露 setHidden；`hidden` 字段只能在读路径保留，
                // 写回 xlsx 时会丢失隐藏标记。sidecar baseline 会继续保住 Univer 侧的 hidden 值。
                // Note: POI 5.2.5 exposes no setHidden on XSSFName; hidden is read-only in round-trip.
            } catch (Exception ignored) {
                // best-effort：单条名称失败不阻塞其他写入
            }
        }
    }
}
