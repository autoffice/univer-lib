package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Base64;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;

/**
 * 图片转换器：在 POI XSSFSheet 的浮动图片与 Univer 的 {@code SHEET_DRAWING_PLUGIN}
 * 插件数据（{@code ISheetImage}）之间双向转换。
 *
 * <p>SHEET_DRAWING_PLUGIN 按 sheet 组织成 {@code {subUnitId: {data: {[drawingId]: ISheetImage}, order: []}}}。
 * 本类只处理单 sheet 级，WorkbookConverter 负责组装 sheet-map。
 *
 * <p>Picture bridge between POI XSSFPicture and Univer's {@code SHEET_DRAWING_PLUGIN} payload.
 * Handles floating over-grid images; BASE64 source is embedded inline so snapshots are self-contained.
 */
public final class PictureConverter {

    /** 1 pixel = 9525 EMU at 96 DPI（OOXML DrawingML 约定）。 */
    private static final int EMU_PER_PIXEL = 9525;
    /** Univer drawingType 枚举：图片。 */
    private static final int DRAWING_TYPE_IMAGE = 0;

    private PictureConverter() {}

    // ============================================================
    // 读路径 / Read path: POI -> Univer {data:{id->ISheetImage}, order:[id...]}
    // ============================================================

    /**
     * 读取 sheet 的全部图片，返回 {@code {data: {drawingId: ISheetImage}, order: [drawingId]}}。
     * unitId 需由 WorkbookConverter 写入（调用方知道它是谁）。
     * 无图片时返回空 ObjectNode（{@code data} 为 {}），永不返回 null。
     */
    public static ObjectNode readSheetPictures(XSSFSheet sheet, ObjectMapper mapper,
                                               String unitId, String subUnitId) {
        ObjectNode out = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        ArrayNode order = mapper.createArrayNode();
        out.set("data", data);
        out.set("order", order);
        if (sheet == null) {
            return out;
        }
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return out;
        }
        List<XSSFShape> shapes = drawing.getShapes();
        if (shapes == null || shapes.isEmpty()) {
            return out;
        }
        for (XSSFShape shape : shapes) {
            if (!(shape instanceof XSSFPicture)) {
                continue;
            }
            XSSFPicture pic = (XSSFPicture) shape;
            XSSFPictureData pd = pic.getPictureData();
            if (pd == null) {
                continue;
            }
            String drawingId = "drawing-" + UUID.randomUUID();
            ObjectNode item = buildSheetImage(pic, pd, unitId, subUnitId, drawingId, mapper);
            data.set(drawingId, item);
            order.add(drawingId);
        }
        return out;
    }

    private static ObjectNode buildSheetImage(XSSFPicture pic, XSSFPictureData pd,
                                              String unitId, String subUnitId, String drawingId,
                                              ObjectMapper mapper) {
        ObjectNode item = mapper.createObjectNode();
        item.put("unitId", unitId);
        item.put("subUnitId", subUnitId);
        item.put("drawingId", drawingId);
        item.put("drawingType", DRAWING_TYPE_IMAGE);
        item.put("imageSourceType", "BASE64");
        String mime = pd.getMimeType();
        if (mime == null || mime.isEmpty()) {
            mime = "image/png";
        }
        String base64 = Base64.getEncoder().encodeToString(pd.getData());
        item.put("source", "data:" + mime + ";base64," + base64);

        // sheetTransform：from / to （四点表示浮动锚定）
        ClientAnchor a = pic.getClientAnchor();
        ObjectNode sheetTransform = mapper.createObjectNode();
        sheetTransform.set("from", overGridPoint(mapper, a.getCol1(), a.getDx1(), a.getRow1(), a.getDy1()));
        sheetTransform.set("to", overGridPoint(mapper, a.getCol2(), a.getDx2(), a.getRow2(), a.getDy2()));
        item.set("sheetTransform", sheetTransform);

        // transform：Univer 渲染用的绝对 box（像素）。大致估算：宽度 = to.col+colOff 的像素 - from 的像素。
        // POI 5.2 未暴露可靠的列宽 -> px 换算（依赖 DPI 与字体），这里只给 top/left 的 offset 与图片内在尺寸；
        // Univer 若拿不到 transform，会根据 sheetTransform + sheet 当前列宽重算，通常足够。
        // transform (optional): omitting it forces Univer to compute from sheetTransform and the
        // current sheet metrics, which is the correct behavior for round-trips.
        return item;
    }

    private static ObjectNode overGridPoint(ObjectMapper mapper, int col, int dxEmu, int row, int dyEmu) {
        ObjectNode n = mapper.createObjectNode();
        n.put("column", col);
        n.put("columnOffset", emuToPx(dxEmu));
        n.put("row", row);
        n.put("rowOffset", emuToPx(dyEmu));
        return n;
    }

    // ============================================================
    // 写路径 / Write path: Univer ISheetImage list -> POI
    // ============================================================

    /**
     * 把 Univer 图片数据写入 sheet。
     * 输入节点形如 {@code {data: {[drawingId]: ISheetImage}, order: [drawingId]}}。
     */
    public static void writeSheetPictures(XSSFSheet sheet, JsonNode sheetPayload) {
        if (sheet == null || sheetPayload == null || !sheetPayload.isObject()) {
            return;
        }
        JsonNode data = sheetPayload.path("data");
        if (!data.isObject() || data.size() == 0) {
            return;
        }
        JsonNode orderNode = sheetPayload.path("order");

        XSSFWorkbook wb = sheet.getWorkbook();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 按 order 顺序处理，无 order 时按 data 自身顺序
        if (orderNode.isArray() && orderNode.size() > 0) {
            for (JsonNode id : orderNode) {
                JsonNode item = data.get(id.asText());
                if (item != null && item.isObject()) {
                    writePicture(sheet, wb, drawing, item);
                }
            }
        } else {
            Iterator<java.util.Map.Entry<String, JsonNode>> it = data.fields();
            while (it.hasNext()) {
                JsonNode item = it.next().getValue();
                if (item != null && item.isObject()) {
                    writePicture(sheet, wb, drawing, item);
                }
            }
        }
    }

    private static void writePicture(XSSFSheet sheet, XSSFWorkbook wb, XSSFDrawing drawing, JsonNode item) {
        // drawingType 非图片直接跳过（Univer 的 SmartArt / shape 等暂不支持）
        int drawingType = item.path("drawingType").asInt(DRAWING_TYPE_IMAGE);
        if (drawingType != DRAWING_TYPE_IMAGE) {
            return;
        }
        String source = item.path("source").asText("");
        String sourceType = item.path("imageSourceType").asText("BASE64");
        byte[] bytes = decodeImage(source, sourceType);
        if (bytes == null || bytes.length == 0) {
            return;
        }
        int poiType = detectPictureType(source, bytes);
        int picIndex = wb.addPicture(bytes, poiType);

        JsonNode st = item.path("sheetTransform");
        JsonNode from = st.path("from");
        JsonNode to = st.path("to");
        XSSFClientAnchor anchor = new XSSFClientAnchor(
                pxToEmu(from.path("columnOffset").asInt(0)),
                pxToEmu(from.path("rowOffset").asInt(0)),
                pxToEmu(to.path("columnOffset").asInt(0)),
                pxToEmu(to.path("rowOffset").asInt(0)),
                from.path("column").asInt(0),
                from.path("row").asInt(0),
                to.path("column").asInt(0),
                to.path("row").asInt(0));
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        drawing.createPicture(anchor, picIndex);
    }

    /** BASE64 data URI / 纯 base64 / URL 三种情况的 bytes 解码。URL 这里不支持下载，跳过。 */
    private static byte[] decodeImage(String source, String sourceType) {
        if (source == null || source.isEmpty()) {
            return new byte[0];
        }
        try {
            if ("BASE64".equalsIgnoreCase(sourceType) || source.startsWith("data:")) {
                int comma = source.indexOf(',');
                String payload = comma >= 0 ? source.substring(comma + 1) : source;
                return Base64.getDecoder().decode(payload);
            }
            // URL / UUID 场景：没有本地字节可写，返回空，跳过这张图。
            // 后续可接 HTTP 下载或 Univer image hosting 解析。
            return new byte[0];
        } catch (IllegalArgumentException e) {
            return new byte[0];
        }
    }

    /** 从 data URI 的 MIME 或 magic bytes 推断 POI 的 picture type 常量。 */
    private static int detectPictureType(String source, byte[] bytes) {
        String mime = null;
        if (source != null && source.startsWith("data:")) {
            int end = source.indexOf(';');
            if (end > 5) {
                mime = source.substring(5, end).toLowerCase();
            }
        }
        if (mime != null) {
            switch (mime) {
                case "image/jpeg":
                case "image/jpg":
                    return Workbook.PICTURE_TYPE_JPEG;
                case "image/png":
                    return Workbook.PICTURE_TYPE_PNG;
                case "image/dib":
                case "image/bmp":
                    return Workbook.PICTURE_TYPE_DIB;
                case "image/wmf":
                    return Workbook.PICTURE_TYPE_WMF;
                case "image/emf":
                    return Workbook.PICTURE_TYPE_EMF;
                case "image/pict":
                    return Workbook.PICTURE_TYPE_PICT;
                default:
                    break;
            }
        }
        // PNG magic: 89 50 4E 47
        if (bytes != null && bytes.length >= 4
                && (bytes[0] & 0xff) == 0x89 && (bytes[1] & 0xff) == 0x50
                && (bytes[2] & 0xff) == 0x4E && (bytes[3] & 0xff) == 0x47) {
            return Workbook.PICTURE_TYPE_PNG;
        }
        // JPEG magic: FF D8
        if (bytes != null && bytes.length >= 2
                && (bytes[0] & 0xff) == 0xFF && (bytes[1] & 0xff) == 0xD8) {
            return Workbook.PICTURE_TYPE_JPEG;
        }
        return Workbook.PICTURE_TYPE_PNG;
    }

    private static int emuToPx(int emu) {
        return (int) Math.round((double) emu / EMU_PER_PIXEL);
    }

    private static int pxToEmu(int px) {
        return px * EMU_PER_PIXEL;
    }
}
