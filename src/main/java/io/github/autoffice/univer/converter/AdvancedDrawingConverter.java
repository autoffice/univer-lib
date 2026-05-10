package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * 高级绘图转换器：图表、形状、迷你图。
 * Advanced drawing converter for charts, non-image shapes, and sparklines.
 *
 * <p>设计：图片走 {@link PictureConverter}（不变），其他三类的可逆性差异较大：
 * <ul>
 *   <li>图表：以 {@code SHEET_CHART_PLUGIN} 资源保留原始 {@code CTChartSpace} XML；
 *       写时把 XML 解析回 {@code CTChartSpace} 再覆盖 POI 创建的空 chart 容器。</li>
 *   <li>形状：以 {@code SHEET_SHAPE_PLUGIN} 资源记录 {@code shapeType / kind / sheetTransform / text}，
 *       足以让 simpleShape / connector 在 xlsx 端可见，并能按枚举名稳定 round-trip。</li>
 *   <li>迷你图：POI 没有高级 API，作为保守方案保留 {@code worksheet.extLst} 整段 XML，
 *       让外部 Excel 仍能识别原 sparkline 定义。</li>
 * </ul>
 *
 * <p>Image drawings continue to flow through {@link PictureConverter}.
 * The other categories use a conservative best-effort scheme:
 * charts are preserved as raw {@code CTChartSpace} XML, shapes are tracked by
 * their POI shape type plus anchor, sparklines are kept as the raw
 * {@code worksheet.extLst} XML so external Excel still renders the trend.
 */
public final class AdvancedDrawingConverter {

    /** 1 像素对应的 EMU 数（96 DPI）/ EMU per pixel at 96 DPI. */
    private static final int EMU_PER_PIXEL = 9525;

    private AdvancedDrawingConverter() {}

    // ============================================================
    // Charts
    // ============================================================

    /**
     * 读 chart payload，{@code {data: {chartId: ChartItem}, order: [chartId]}}。
     * Read charts as a {data, order} envelope similar to drawings.
     */
    public static ObjectNode readSheetCharts(XSSFSheet sheet, ObjectMapper mapper,
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
        List<XSSFChart> charts = drawing.getCharts();
        if (charts == null || charts.isEmpty()) {
            return out;
        }
        for (XSSFChart chart : charts) {
            if (chart == null) {
                continue;
            }
            String chartId = "chart-" + UUID.randomUUID();
            ObjectNode item = mapper.createObjectNode();
            item.put("unitId", unitId);
            item.put("subUnitId", subUnitId);
            item.put("chartId", chartId);
            item.put("rawXml", chart.getCTChartSpace().xmlText());
            if (chart.getTitleText() != null) {
                item.put("title", chart.getTitleText().getString());
            }
            XSSFGraphicFrame frame = chart.getGraphicFrame();
            if (frame != null) {
                XSSFAnchor anchor = frame.getAnchor();
                if (anchor instanceof XSSFClientAnchor) {
                    item.set("sheetTransform", anchorToNode(mapper, (XSSFClientAnchor) anchor));
                }
            }
            data.set(chartId, item);
            order.add(chartId);
        }
        return out;
    }

    /**
     * 写 chart payload。先用 POI 创建空 chart 占位，再用 sidecar 中保存的 {@code rawXml}
     * 覆盖其 {@code CTChartSpace}。解析失败时保留空 chart，避免破坏整个 xlsx。
     * Write charts: create an empty chart container then overwrite its
     * {@code CTChartSpace} with the preserved XML.
     */
    public static void writeSheetCharts(XSSFSheet sheet, JsonNode sheetPayload) {
        if (sheet == null || sheetPayload == null || !sheetPayload.isObject()) {
            return;
        }
        JsonNode data = sheetPayload.path("data");
        if (!data.isObject() || data.size() == 0) {
            return;
        }
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        for (JsonNode item : orderedItems(data, sheetPayload.path("order"))) {
            if (item == null || !item.isObject()) {
                continue;
            }
            String rawXml = item.path("rawXml").asText("");
            if (rawXml.isEmpty()) {
                continue;
            }
            XSSFClientAnchor anchor = anchorFromNode(item.path("sheetTransform"));
            XSSFChart chart = drawing.createChart(anchor);
            try {
                CTChartSpace chartSpace = CTChartSpace.Factory.parse(rawXml);
                chart.getCTChartSpace().set(chartSpace);
            } catch (Exception ignored) {
                // 保守降级：保留空 chart 容器。
            }
        }
    }

    // ============================================================
    // Sparkline (worksheet extLst raw XML)
    // ============================================================

    /**
     * 读取 sheet 的迷你图占位（{@code worksheet.extLst} 整段 XML）。
     * Read worksheet-level sparkline payload as raw extLst XML.
     */
    public static ObjectNode readSheetSparkline(XSSFSheet sheet, ObjectMapper mapper) {
        ObjectNode out = mapper.createObjectNode();
        if (sheet == null) {
            return out;
        }
        CTWorksheet worksheet = sheet.getCTWorksheet();
        if (worksheet != null && worksheet.isSetExtLst()) {
            out.put("extLstXml", worksheet.getExtLst().xmlText());
        }
        return out;
    }

    /**
     * 写迷你图：把 sidecar 中保存的 extLst XML 还原回 {@code CTWorksheet}。
     * Write sparkline payload by restoring the preserved extLst XML.
     */
    public static void writeSheetSparkline(XSSFSheet sheet, JsonNode sheetPayload) {
        if (sheet == null || sheetPayload == null || !sheetPayload.isObject()) {
            return;
        }
        String extLstXml = sheetPayload.path("extLstXml").asText("");
        if (extLstXml.isEmpty()) {
            return;
        }
        try {
            CTExtensionList extLst = CTExtensionList.Factory.parse(extLstXml);
            sheet.getCTWorksheet().setExtLst(extLst);
        } catch (Exception ignored) {
            // sparkline 只做保守保留。
        }
    }

    // ============================================================
    // Non-image shapes (simpleShape + connector)
    // ============================================================

    /**
     * 读取 sheet 上非图片、非图表的形状，按 {@code {data, order}} 结构返回。
     * Read non-image, non-chart shapes (simpleShape / connector).
     */
    public static ObjectNode readSheetShapes(XSSFSheet sheet, ObjectMapper mapper,
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
            if (shape == null) {
                continue;
            }
            if (shape instanceof XSSFPicture || shape instanceof XSSFGraphicFrame) {
                continue;
            }
            if (!(shape instanceof XSSFSimpleShape || shape instanceof XSSFConnector)) {
                continue;
            }
            String shapeId = "shape-" + UUID.randomUUID();
            ObjectNode item = mapper.createObjectNode();
            item.put("unitId", unitId);
            item.put("subUnitId", subUnitId);
            item.put("shapeId", shapeId);
            item.put("kind", shape instanceof XSSFConnector ? "connector" : "shape");
            int shapeType = shape instanceof XSSFConnector
                    ? ((XSSFConnector) shape).getShapeType()
                    : ((XSSFSimpleShape) shape).getShapeType();
            item.put("shapeType", shapeTypeNameOf(shapeType));
            if (shape instanceof XSSFSimpleShape) {
                String text = ((XSSFSimpleShape) shape).getText();
                if (text != null && !text.isEmpty()) {
                    item.put("text", text);
                }
            }
            XSSFAnchor anchor = shape.getAnchor();
            if (anchor instanceof XSSFClientAnchor) {
                item.set("sheetTransform", anchorToNode(mapper, (XSSFClientAnchor) anchor));
            }
            data.set(shapeId, item);
            order.add(shapeId);
        }
        return out;
    }

    /**
     * 写形状：simpleShape 与 connector 共享同一种锚点格式；为避免连线引用未创建的形状，
     * 先建出全部 simpleShape，再处理 connector。
     * Write shapes by creating simpleShapes first, then connectors so they
     * can reference shapes created earlier.
     */
    public static void writeSheetShapes(XSSFSheet sheet, JsonNode sheetPayload) {
        if (sheet == null || sheetPayload == null || !sheetPayload.isObject()) {
            return;
        }
        JsonNode data = sheetPayload.path("data");
        if (!data.isObject() || data.size() == 0) {
            return;
        }
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        List<JsonNode> items = orderedItems(data, sheetPayload.path("order"));
        Map<String, XSSFSimpleShape> simpleShapes = new LinkedHashMap<>();
        for (JsonNode item : items) {
            if (item == null || !item.isObject()) {
                continue;
            }
            if ("connector".equals(item.path("kind").asText("shape"))) {
                continue;
            }
            XSSFSimpleShape shape = createSimpleShape(drawing, item);
            String id = item.path("shapeId").asText("");
            if (!id.isEmpty()) {
                simpleShapes.put(id, shape);
            }
        }
        for (JsonNode item : items) {
            if (item == null || !item.isObject()) {
                continue;
            }
            if (!"connector".equals(item.path("kind").asText("shape"))) {
                continue;
            }
            createConnector(drawing, item);
        }
        // simpleShapes 仅用于将来扩展（如连接锚点引用）；当前未直接消费。
        // Reserved for future enrichment such as connector site lookup.
        if (simpleShapes.isEmpty()) {
            // no-op
        }
    }

    private static XSSFSimpleShape createSimpleShape(XSSFDrawing drawing, JsonNode item) {
        XSSFClientAnchor anchor = anchorFromNode(item.path("sheetTransform"));
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(shapeTypeCodeOf(item.path("shapeType").asText("rect")));
        String text = item.path("text").asText(null);
        if (text != null) {
            shape.setText(text);
        }
        return shape;
    }

    private static void createConnector(XSSFDrawing drawing, JsonNode item) {
        XSSFClientAnchor anchor = anchorFromNode(item.path("sheetTransform"));
        XSSFConnector connector = drawing.createConnector(anchor);
        connector.setShapeType(shapeTypeCodeOf(item.path("shapeType").asText("straightConnector1")));
    }

    // ============================================================
    // Helpers
    // ============================================================

    private static List<JsonNode> orderedItems(JsonNode data, JsonNode orderNode) {
        List<JsonNode> result = new ArrayList<>();
        if (orderNode != null && orderNode.isArray() && orderNode.size() > 0) {
            for (JsonNode id : orderNode) {
                JsonNode item = data.get(id.asText());
                if (item != null) {
                    result.add(item);
                }
            }
            return result;
        }
        Iterator<Map.Entry<String, JsonNode>> it = data.fields();
        while (it.hasNext()) {
            result.add(it.next().getValue());
        }
        return result;
    }

    private static ObjectNode anchorToNode(ObjectMapper mapper, XSSFClientAnchor anchor) {
        ObjectNode out = mapper.createObjectNode();
        ObjectNode from = mapper.createObjectNode();
        from.put("column", anchor.getCol1());
        from.put("columnOffset", emuToPx(anchor.getDx1()));
        from.put("row", anchor.getRow1());
        from.put("rowOffset", emuToPx(anchor.getDy1()));
        ObjectNode to = mapper.createObjectNode();
        to.put("column", anchor.getCol2());
        to.put("columnOffset", emuToPx(anchor.getDx2()));
        to.put("row", anchor.getRow2());
        to.put("rowOffset", emuToPx(anchor.getDy2()));
        out.set("from", from);
        out.set("to", to);
        return out;
    }

    private static XSSFClientAnchor anchorFromNode(JsonNode node) {
        JsonNode from = node.path("from");
        JsonNode to = node.path("to");
        return new XSSFClientAnchor(
                pxToEmu(from.path("columnOffset").asInt(0)),
                pxToEmu(from.path("rowOffset").asInt(0)),
                pxToEmu(to.path("columnOffset").asInt(0)),
                pxToEmu(to.path("rowOffset").asInt(0)),
                from.path("column").asInt(0),
                from.path("row").asInt(0),
                to.path("column").asInt(0),
                to.path("row").asInt(0));
    }

    private static String shapeTypeNameOf(int shapeType) {
        switch (shapeType) {
            case ShapeTypes.RECT:
                return "rect";
            case ShapeTypes.ROUND_RECT:
                return "roundRect";
            case ShapeTypes.DIAMOND:
                return "diamond";
            case ShapeTypes.SMILEY_FACE:
                return "smileyFace";
            case ShapeTypes.FLOW_CHART_PROCESS:
                return "flowChartProcess";
            case ShapeTypes.FLOW_CHART_DECISION:
                return "flowChartDecision";
            case ShapeTypes.FLOW_CHART_TERMINATOR:
                return "flowChartTerminator";
            case ShapeTypes.FLOW_CHART_DOCUMENT:
                return "flowChartDocument";
            case ShapeTypes.STRAIGHT_CONNECTOR_1:
                return "straightConnector1";
            case ShapeTypes.BENT_CONNECTOR_2:
                return "bentConnector2";
            case ShapeTypes.BENT_CONNECTOR_3:
                return "bentConnector3";
            case ShapeTypes.CURVED_CONNECTOR_2:
                return "curvedConnector2";
            case ShapeTypes.CURVED_CONNECTOR_3:
                return "curvedConnector3";
            case ShapeTypes.DOWN_ARROW:
                return "downArrow";
            default:
                return "rect";
        }
    }

    private static int shapeTypeCodeOf(String shapeType) {
        if (shapeType == null || shapeType.isEmpty()) {
            return ShapeTypes.RECT;
        }
        switch (shapeType) {
            case "roundRect":
                return ShapeTypes.ROUND_RECT;
            case "diamond":
                return ShapeTypes.DIAMOND;
            case "smileyFace":
                return ShapeTypes.SMILEY_FACE;
            case "flowChartProcess":
                return ShapeTypes.FLOW_CHART_PROCESS;
            case "flowChartDecision":
                return ShapeTypes.FLOW_CHART_DECISION;
            case "flowChartTerminator":
                return ShapeTypes.FLOW_CHART_TERMINATOR;
            case "flowChartDocument":
                return ShapeTypes.FLOW_CHART_DOCUMENT;
            case "straightConnector1":
                return ShapeTypes.STRAIGHT_CONNECTOR_1;
            case "bentConnector2":
                return ShapeTypes.BENT_CONNECTOR_2;
            case "bentConnector3":
                return ShapeTypes.BENT_CONNECTOR_3;
            case "curvedConnector2":
                return ShapeTypes.CURVED_CONNECTOR_2;
            case "curvedConnector3":
                return ShapeTypes.CURVED_CONNECTOR_3;
            case "downArrow":
                return ShapeTypes.DOWN_ARROW;
            case "rect":
            default:
                return ShapeTypes.RECT;
        }
    }

    private static int emuToPx(int emu) {
        return (int) Math.round((double) emu / EMU_PER_PIXEL);
    }

    private static int pxToEmu(int px) {
        return px * EMU_PER_PIXEL;
    }
}
