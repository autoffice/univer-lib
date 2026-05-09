package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;
import java.util.Map;

/**
 * 单元格批注转换器：在 POI XSSFSheet 的老式 comment（xl/comments*.xml）与
 * Univer 的 {@code SHEET_NOTE_PLUGIN} 资源之间双向转换。
 *
 * <p>Univer 侧的 per-sheet 批注表结构：
 * <pre>
 * {
 *   "<rowIdx>": {
 *     "<colIdx>": { "note": "...", "width": 160, "height": 100, "show": false }
 *   }
 * }
 * </pre>
 *
 * <p>Cell comment bridge between POI XSSFSheet's legacy comments and Univer's
 * {@code SHEET_NOTE_PLUGIN} plugin payload (ISheetNote).
 */
public final class CommentConverter {

    /** ISheetNote 的缺省宽度（px），与 Facade createOrUpdateNote 文档示例一致。 */
    private static final int DEFAULT_WIDTH_PX = 160;
    /** ISheetNote 的缺省高度（px）。 */
    private static final int DEFAULT_HEIGHT_PX = 100;

    private CommentConverter() {}

    // ============================================================
    // 读路径 / Read path: POI -> Univer {row:{col:note}} JSON
    // ============================================================

    /**
     * 读取 sheet 的全部单元格批注，按 row/col 嵌套写入 ObjectNode。
     * Read all cell comments from a sheet into a nested {row: {col: ISheetNote}} JSON.
     *
     * @return 嵌套对象；无批注时返回空 ObjectNode，永不返回 null。
     */
    public static ObjectNode readSheetComments(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (sheet == null || !sheet.hasComments()) {
            return out;
        }
        Map<CellAddress, XSSFComment> comments = sheet.getCellComments();
        if (comments == null || comments.isEmpty()) {
            return out;
        }
        for (Map.Entry<CellAddress, XSSFComment> e : comments.entrySet()) {
            XSSFComment c = e.getValue();
            if (c == null) {
                continue;
            }
            int row = c.getRow();
            int col = c.getColumn();
            RichTextString rts = c.getString();
            String note = rts == null ? "" : rts.getString();
            if (note == null) {
                note = "";
            }
            ObjectNode cell = mapper.createObjectNode();
            cell.put("note", note);
            cell.put("width", DEFAULT_WIDTH_PX);
            cell.put("height", DEFAULT_HEIGHT_PX);
            cell.put("show", c.isVisible());
            String rowKey = String.valueOf(row);
            ObjectNode rowNode = (ObjectNode) out.get(rowKey);
            if (rowNode == null) {
                rowNode = mapper.createObjectNode();
                out.set(rowKey, rowNode);
            }
            rowNode.set(String.valueOf(col), cell);
        }
        return out;
    }

    // ============================================================
    // 写路径 / Write path: Univer {row:{col:note}} JSON -> POI
    // ============================================================

    /**
     * 把 Univer 批注 JSON 写入 sheet。键可以是字符串或整数（Jackson 的字段名始终是字符串）。
     * Apply a Univer note JSON map onto the sheet as legacy xlsx comments.
     */
    public static void writeSheetComments(XSSFSheet sheet, JsonNode rowMap) {
        if (sheet == null || rowMap == null || !rowMap.isObject() || rowMap.size() == 0) {
            return;
        }
        XSSFWorkbook wb = sheet.getWorkbook();
        CreationHelper helper = wb.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        Iterator<Map.Entry<String, JsonNode>> rows = rowMap.fields();
        while (rows.hasNext()) {
            Map.Entry<String, JsonNode> rowEntry = rows.next();
            int row = parseIntOrSkip(rowEntry.getKey());
            if (row < 0) {
                continue;
            }
            JsonNode colMap = rowEntry.getValue();
            if (colMap == null || !colMap.isObject()) {
                continue;
            }
            Iterator<Map.Entry<String, JsonNode>> cols = colMap.fields();
            while (cols.hasNext()) {
                Map.Entry<String, JsonNode> colEntry = cols.next();
                int col = parseIntOrSkip(colEntry.getKey());
                if (col < 0) {
                    continue;
                }
                JsonNode noteNode = colEntry.getValue();
                if (noteNode == null || !noteNode.isObject()) {
                    continue;
                }
                String note = noteNode.path("note").asText("");
                boolean show = noteNode.path("show").asBoolean(false);
                writeComment(sheet, drawing, helper, row, col, note, show);
            }
        }
    }

    private static void writeComment(XSSFSheet sheet, Drawing<?> drawing, CreationHelper helper,
                                     int row, int col, String note, boolean show) {
        // 目标单元格必须存在，comment 才有归属 / the target cell must exist to anchor the comment
        if (sheet.getRow(row) == null) {
            sheet.createRow(row);
        }
        if (sheet.getRow(row).getCell(col) == null) {
            sheet.getRow(row).createCell(col);
        }
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(col);
        anchor.setRow1(row);
        // Excel 默认批注显示框大约 2 列 × 4 行 / Excel's default comment box spans ~2 cols × 4 rows
        anchor.setCol2(col + 2);
        anchor.setRow2(row + 4);
        if (anchor instanceof XSSFClientAnchor) {
            ((XSSFClientAnchor) anchor).setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setAddress(new CellAddress(row, col));
        comment.setString(new XSSFRichTextString(note == null ? "" : note));
        comment.setVisible(show);
    }

    private static int parseIntOrSkip(String s) {
        if (s == null || s.isEmpty()) {
            return -1;
        }
        try {
            return Integer.parseInt(s);
        } catch (NumberFormatException e) {
            return -1;
        }
    }
}
