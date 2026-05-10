package io.github.autoffice.univer.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtension;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 高级 drawing 转换器测试：图表、形状、迷你图。
 * Roundtrip tests for charts, shapes, and sparklines.
 */
class AdvancedDrawingConverterTest {
    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_roundtrip_chart_via_raw_xml() throws Exception {
        ObjectNode payload;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("A");
            sh.getRow(0).createCell(1).setCellValue(1);
            sh.createRow(1).createCell(0).setCellValue("B");
            sh.getRow(1).createCell(1).setCellValue(2);
            XSSFDrawing drawing = sh.createDrawingPatriarch();
            XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, 1, 8, 15));
            XDDFCategoryAxis bottom = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis left = chart.createValueAxis(AxisPosition.LEFT);
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottom, left);
            data.addSeries(
                    XDDFDataSourcesFactory.fromStringCellRange(sh, new CellRangeAddress(0, 1, 0, 0)),
                    XDDFDataSourcesFactory.fromNumericCellRange(sh, new CellRangeAddress(0, 1, 1, 1)));
            chart.plot(data);
            chart.setTitleText("Demo");
            payload = AdvancedDrawingConverter.readSheetCharts(sh, mapper, "u-1", "S1");
        }
        assertThat(payload.path("data").size()).isEqualTo(1);
        JsonNode item = payload.path("data").fields().next().getValue();
        assertThat(item.path("rawXml").asText()).contains("chart");
        assertThat(item.path("title").asText()).isEqualTo("Demo");

        byte[] roundTripped = writeReadBack(payload, AdvancedDrawingConverter::writeSheetCharts);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(roundTripped))) {
            XSSFSheet sh = wb.getSheetAt(0);
            List<XSSFChart> charts = sh.getDrawingPatriarch().getCharts();
            assertThat(charts).hasSize(1);
            assertThat(charts.get(0).getTitleText().getString()).isEqualTo("Demo");
        }
    }

    @Test
    void should_roundtrip_simple_shape_with_text() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        ObjectNode shape = shapeNode("shape-1", "shape", "roundRect", "Hello",
                1, 0, 1, 0, 4, 0, 4, 0);
        data.set("shape-1", shape);
        payload.set("data", data);
        ArrayNode order = mapper.createArrayNode();
        order.add("shape-1");
        payload.set("order", order);

        byte[] xlsx = writeReadBack(payload, AdvancedDrawingConverter::writeSheetShapes);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            XSSFSheet sh = wb.getSheetAt(0);
            List<XSSFShape> shapes = sh.getDrawingPatriarch().getShapes();
            assertThat(shapes).filteredOn(s -> s instanceof XSSFSimpleShape).hasSize(1);
            XSSFSimpleShape simple = (XSSFSimpleShape) shapes.stream()
                    .filter(s -> s instanceof XSSFSimpleShape).findFirst().orElseThrow(AssertionError::new);
            assertThat(simple.getShapeType()).isEqualTo(ShapeTypes.ROUND_RECT);
            assertThat(simple.getText()).contains("Hello");

            ObjectNode reread = AdvancedDrawingConverter.readSheetShapes(sh, mapper, "u-1", "S1");
            assertThat(reread.path("data").size()).isEqualTo(1);
            JsonNode rt = reread.path("data").fields().next().getValue();
            assertThat(rt.path("kind").asText()).isEqualTo("shape");
            assertThat(rt.path("shapeType").asText()).isEqualTo("roundRect");
            assertThat(rt.path("text").asText()).contains("Hello");
            assertThat(rt.path("sheetTransform").path("from").path("column").asInt()).isEqualTo(1);
            assertThat(rt.path("sheetTransform").path("to").path("column").asInt()).isEqualTo(4);
        }
    }

    @Test
    void should_roundtrip_connector_shape() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        ObjectNode data = mapper.createObjectNode();
        ObjectNode connector = shapeNode("conn-1", "connector", "straightConnector1", null,
                0, 0, 0, 0, 5, 0, 5, 0);
        data.set("conn-1", connector);
        payload.set("data", data);
        ArrayNode order = mapper.createArrayNode();
        order.add("conn-1");
        payload.set("order", order);

        byte[] xlsx = writeReadBack(payload, AdvancedDrawingConverter::writeSheetShapes);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            XSSFSheet sh = wb.getSheetAt(0);
            List<XSSFShape> shapes = sh.getDrawingPatriarch().getShapes();
            assertThat(shapes).filteredOn(s -> s instanceof XSSFConnector).hasSize(1);
            XSSFConnector cn = (XSSFConnector) shapes.stream()
                    .filter(s -> s instanceof XSSFConnector).findFirst().orElseThrow(AssertionError::new);
            assertThat(cn.getShapeType()).isEqualTo(ShapeTypes.STRAIGHT_CONNECTOR_1);

            ObjectNode reread = AdvancedDrawingConverter.readSheetShapes(sh, mapper, "u-1", "S1");
            assertThat(reread.path("data").size()).isEqualTo(1);
            JsonNode rt = reread.path("data").fields().next().getValue();
            assertThat(rt.path("kind").asText()).isEqualTo("connector");
            assertThat(rt.path("shapeType").asText()).isEqualTo("straightConnector1");
        }
    }

    @Test
    void should_skip_when_shape_payload_is_empty() throws Exception {
        ObjectNode payload = mapper.createObjectNode();
        payload.set("data", mapper.createObjectNode());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            AdvancedDrawingConverter.writeSheetShapes(sh, payload);
            assertThat(sh.getDrawingPatriarch()).isNull();
        }
    }

    @Test
    void should_preserve_sparkline_extLst_xml() throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            CTExtensionList extLst = sh.getCTWorksheet().addNewExtLst();
            CTExtension ext = extLst.addNewExt();
            ext.setUri("{05C60535-1F16-4fd2-B633-F4F36F0B0DEC}");
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        ObjectNode captured;
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            captured = AdvancedDrawingConverter.readSheetSparkline(wb.getSheetAt(0), mapper);
        }
        assertThat(captured.path("extLstXml").asText()).contains("05C60535-1F16-4fd2-B633-F4F36F0B0DEC");

        byte[] rewritten;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            AdvancedDrawingConverter.writeSheetSparkline(sh, captured);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                rewritten = out.toByteArray();
            }
        }
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(rewritten))) {
            XSSFSheet sh = wb.getSheetAt(0);
            assertThat(sh.getCTWorksheet().isSetExtLst()).isTrue();
            CTExtensionList extLst = sh.getCTWorksheet().getExtLst();
            assertThat(extLst.getExtArray()).extracting(CTExtension::getUri)
                    .contains("{05C60535-1F16-4fd2-B633-F4F36F0B0DEC}");
        }
    }

    @FunctionalInterface
    private interface Writer {
        void apply(XSSFSheet sheet, JsonNode payload);
    }

    private byte[] writeReadBack(ObjectNode payload, Writer writer) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S1");
            sh.createRow(0).createCell(0).setCellValue("A");
            sh.getRow(0).createCell(1).setCellValue(1);
            sh.createRow(1).createCell(0).setCellValue("B");
            sh.getRow(1).createCell(1).setCellValue(2);
            writer.apply(sh, payload);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                return out.toByteArray();
            }
        }
    }

    private ObjectNode shapeNode(String id, String kind, String shapeType, String text,
                                 int fromCol, int fromColOff, int fromRow, int fromRowOff,
                                 int toCol, int toColOff, int toRow, int toRowOff) {
        ObjectNode n = mapper.createObjectNode();
        n.put("unitId", "u-1");
        n.put("subUnitId", "S1");
        n.put("shapeId", id);
        n.put("kind", kind);
        n.put("shapeType", shapeType);
        if (text != null) {
            n.put("text", text);
        }
        ObjectNode from = mapper.createObjectNode();
        from.put("column", fromCol);
        from.put("columnOffset", fromColOff);
        from.put("row", fromRow);
        from.put("rowOffset", fromRowOff);
        ObjectNode to = mapper.createObjectNode();
        to.put("column", toCol);
        to.put("columnOffset", toColOff);
        to.put("row", toRow);
        to.put("rowOffset", toRowOff);
        ObjectNode st = mapper.createObjectNode();
        st.set("from", from);
        st.set("to", to);
        n.set("sheetTransform", st);
        return n;
    }
}
