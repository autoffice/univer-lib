package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.IBorderData;
import io.github.autoffice.univer.model.IBorderStyleData;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IColorStyle;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IFreeze;
import io.github.autoffice.univer.model.INumfmtLocal;
import io.github.autoffice.univer.model.IRange;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextDecoration;
import io.github.autoffice.univer.model.ITextRotation;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 完整 round-trip 集成测试，覆盖多种边界场景。
 * Full round-trip integration tests covering edge cases.
 */
class FullRoundTripTest {

    // ============================================================
    // helpers
    // ============================================================

    private byte[] writeBytes(IWorkbookData src) throws Exception {
        return writeBytes(src, UniverXlsxOptions.defaults());
    }

    private byte[] writeBytes(IWorkbookData src, UniverXlsxOptions opts) throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf, opts);
        return buf.toByteArray();
    }

    private IWorkbookData roundtrip(IWorkbookData src) throws Exception {
        return UniverXlsx.read(new ByteArrayInputStream(writeBytes(src)));
    }

    private IWorksheetData newSheetWithCell(String id, String name, String cellValue) {
        IWorksheetData ws = new IWorksheetData().setId(id).setName(name);
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV(cellValue).setT(CellValueType.STRING));
        ws.getCellData().put(0, row);
        return ws;
    }

    /**
     * 解析 cell.s 为 IStyleData（容忍 String/IStyleData/Map 三种形式）。
     * Resolve cell.s to IStyleData regardless of String/IStyleData/Map form.
     */
    private IStyleData resolveStyle(IWorkbookData wb, ICellData cell) {
        Object raw = cell.getS();
        if (raw == null) {
            return null;
        }
        if (raw instanceof IStyleData) {
            return (IStyleData) raw;
        }
        if (raw instanceof String) {
            return wb.getStyles() == null ? null : wb.getStyles().get((String) raw);
        }
        // Map form (来自 sidecar JSON 反序列化为 Object 时)
        return io.github.autoffice.univer.util.JsonMapper.get().convertValue(raw, IStyleData.class);
    }

    // ============================================================
    // 1. sheetOrder + hidden
    // ============================================================

    @Test
    void should_roundtrip_three_sheets_with_one_hidden_and_disordered_sheetOrder() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2").setLocale("zhCN");
        IWorksheetData s1 = newSheetWithCell("s1", "S1", "a");
        IWorksheetData s2 = newSheetWithCell("s2", "S2", "b").setHidden(BooleanNumber.TRUE);
        IWorksheetData s3 = newSheetWithCell("s3", "S3", "c");
        src.getSheets().put("s1", s1);
        src.getSheets().put("s2", s2);
        src.getSheets().put("s3", s3);
        src.setSheetOrder(Arrays.asList("s3", "s1", "s2"));

        IWorkbookData back = roundtrip(src);

        assertThat(back.getSheetOrder()).containsExactly("s3", "s1", "s2");
        // 严格检查：sheets 键仅包含 sidecar id，无重复影子项 / Only sidecar ids, no shadow name entries.
        assertThat(back.getSheets().keySet()).containsExactlyInAnyOrder("s1", "s2", "s3");
        assertThat(back.getSheets().get("s2").getHidden()).isEqualTo(BooleanNumber.TRUE);
        assertThat(back.getSheets().get("s1").getHidden()).isNotEqualTo(BooleanNumber.TRUE);
        assertThat(back.getSheets().get("s3").getHidden()).isNotEqualTo(BooleanNumber.TRUE);
    }

    // ============================================================
    // 1b. sheet id preserved across roundtrip
    // ============================================================

    @Test
    void should_preserve_sheet_id_across_roundtrip() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = newSheetWithCell("s1", "Visible", "x");
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        assertThat(back.getSheets()).hasSize(1);
        assertThat(back.getSheets()).containsOnlyKeys("s1");
        assertThat(back.getSheets().get("s1").getName()).isEqualTo("Visible");
    }

    // ============================================================
    // 2. full style data
    // ============================================================

    @Test
    void should_roundtrip_full_style_data() throws Exception {
        IStyleData style = new IStyleData()
                .setFf("Arial")
                .setFs(14)
                .setBl(BooleanNumber.TRUE)
                .setIt(BooleanNumber.TRUE)
                .setUl(new ITextDecoration().setS(BooleanNumber.TRUE))
                .setSt(new ITextDecoration().setS(BooleanNumber.TRUE))
                .setBg(new IColorStyle().setRgb("#ffff00"))
                .setBd(new IBorderData().setT(new IBorderStyleData()
                        .setS(1).setCl(new IColorStyle().setRgb("#000000"))))
                .setCl(new IColorStyle().setRgb("#ff0000"))
                .setHt(2)
                .setVt(2)
                .setTb(3)
                .setTr(new ITextRotation().setA(45))
                .setN(new INumfmtLocal().setPattern("yyyy-mm-dd"));

        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("styled").setT(CellValueType.STRING).setS(style));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        ICellData cell = back.getSheets().get("s1").getCellData().get(0).get(0);
        assertThat(cell.getS()).isNotNull();
        IStyleData s = resolveStyle(back, cell);
        assertThat(s).isNotNull();
        assertThat(s.getFf()).isEqualTo("Arial");
        assertThat(s.getFs()).isEqualTo(14);
        assertThat(s.getBl()).isEqualTo(BooleanNumber.TRUE);
        assertThat(s.getIt()).isEqualTo(BooleanNumber.TRUE);
        assertThat(s.getUl()).isNotNull();
        assertThat(s.getUl().getS()).isEqualTo(BooleanNumber.TRUE);
        assertThat(s.getSt()).isNotNull();
        assertThat(s.getSt().getS()).isEqualTo(BooleanNumber.TRUE);
        assertThat(s.getBg()).isNotNull();
        assertThat(s.getBg().getRgb()).isEqualToIgnoringCase("#ffff00");
        assertThat(s.getCl()).isNotNull();
        assertThat(s.getCl().getRgb()).isEqualToIgnoringCase("#ff0000");
        assertThat(s.getBd()).isNotNull();
        assertThat(s.getBd().getT()).isNotNull();
        assertThat(s.getBd().getT().getS()).isEqualTo(1);
        assertThat(s.getHt()).isEqualTo(2);
        assertThat(s.getVt()).isEqualTo(2);
        assertThat(s.getTb()).isEqualTo(3);
        assertThat(s.getTr()).isNotNull();
        assertThat(s.getTr().getA()).isEqualTo(45);
        assertThat(s.getN()).isNotNull();
        assertThat(s.getN().getPattern()).isEqualTo("yyyy-mm-dd");
    }

    // ============================================================
    // 3. shared formula si
    // ============================================================

    @Test
    void should_roundtrip_shared_formula_si() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        // master at (2,3); followers (0,0), (1,1), (1,2)
        int[][] coords = {{0, 0}, {1, 1}, {1, 2}, {2, 3}};
        for (int[] rc : coords) {
            ICellData c = new ICellData().setF("=SUM(A1:B1)").setSi("shared1");
            ws.getCellData()
                    .computeIfAbsent(rc[0], k -> new LinkedHashMap<>())
                    .put(rc[1], c);
        }
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        IWorksheetData backWs = back.getSheets().get("s1");
        String firstSi = null;
        for (int[] rc : coords) {
            ICellData c = backWs.getCellData().get(rc[0]).get(rc[1]);
            assertThat(c).as("cell at %s,%s", rc[0], rc[1]).isNotNull();
            assertThat(c.getSi()).as("si at %s,%s", rc[0], rc[1]).isNotNull();
            if (firstSi == null) {
                firstSi = c.getSi();
            } else {
                assertThat(c.getSi()).isEqualTo(firstSi);
            }
        }
    }

    // ============================================================
    // 4. rich text
    // ============================================================

    @Test
    void should_roundtrip_rich_text() throws Exception {
        IDocumentData doc = new IDocumentData();
        IDocumentData.Body body = new IDocumentData.Body().setDataStream("hello world");
        body.setTextRuns(Arrays.asList(
                new IDocumentData.TextRun().setSt(0).setEd(5)
                        .setTs(new IStyleData().setFf("Arial").setFs(10)
                                .setCl(new IColorStyle().setRgb("#000000"))),
                new IDocumentData.TextRun().setSt(6).setEd(11)
                        .setTs(new IStyleData().setFf("Arial").setFs(14)
                                .setCl(new IColorStyle().setRgb("#ff0000")))
        ));
        doc.setBody(body);

        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setP(doc).setT(CellValueType.STRING));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        ICellData backCell = back.getSheets().get("s1").getCellData().get(0).get(0);
        assertThat(backCell.getP()).isNotNull();
        assertThat(backCell.getP().getBody()).isNotNull();
        assertThat(backCell.getP().getBody().getDataStream()).isEqualTo("hello world");
        assertThat(backCell.getP().getBody().getTextRuns()).isNotNull();
        assertThat(backCell.getP().getBody().getTextRuns().size()).isGreaterThanOrEqualTo(1);
    }

    // ============================================================
    // 5. complex merges
    // ============================================================

    @Test
    void should_roundtrip_complex_merges() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> r = new LinkedHashMap<>();
        r.put(0, new ICellData().setV("x").setT(CellValueType.STRING));
        ws.getCellData().put(0, r);
        ws.setMergeData(Arrays.asList(
                new IRange().setStartRow(0).setStartColumn(0).setEndRow(1).setEndColumn(1),
                new IRange().setStartRow(3).setStartColumn(0).setEndRow(3).setEndColumn(4),
                new IRange().setStartRow(0).setStartColumn(6).setEndRow(4).setEndColumn(6)
        ));
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        List<IRange> merges = back.getSheets().get("s1").getMergeData();
        assertThat(merges).hasSize(3);
    }

    // ============================================================
    // 6. freeze in three forms
    // ============================================================

    @Test
    void should_roundtrip_freeze_in_three_forms() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData h = newSheetWithCell("h", "frozenH", "h")
                .setFreeze(new IFreeze().setXSplit(0).setYSplit(1).setStartRow(1).setStartColumn(0));
        IWorksheetData v = newSheetWithCell("v", "frozenV", "v")
                .setFreeze(new IFreeze().setXSplit(1).setYSplit(0).setStartRow(0).setStartColumn(1));
        IWorksheetData b = newSheetWithCell("b", "frozenBoth", "b")
                .setFreeze(new IFreeze().setXSplit(1).setYSplit(1).setStartRow(1).setStartColumn(1));
        src.getSheets().put("h", h);
        src.getSheets().put("v", v);
        src.getSheets().put("b", b);
        src.setSheetOrder(Arrays.asList("h", "v", "b"));

        IWorkbookData back = roundtrip(src);

        IFreeze fh = back.getSheets().get("h").getFreeze();
        assertThat(fh).isNotNull();
        assertThat(fh.getYSplit()).isEqualTo(1);
        assertThat(fh.getStartRow()).isEqualTo(1);

        IFreeze fv = back.getSheets().get("v").getFreeze();
        assertThat(fv).isNotNull();
        assertThat(fv.getXSplit()).isEqualTo(1);
        assertThat(fv.getStartColumn()).isEqualTo(1);

        IFreeze fb = back.getSheets().get("b").getFreeze();
        assertThat(fb).isNotNull();
        assertThat(fb.getXSplit()).isEqualTo(1);
        assertThat(fb.getYSplit()).isEqualTo(1);
        assertThat(fb.getStartRow()).isEqualTo(1);
        assertThat(fb.getStartColumn()).isEqualTo(1);
    }

    // ============================================================
    // 7. resources
    // ============================================================

    @Test
    void should_roundtrip_resources_field() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = newSheetWithCell("s1", "S1", "x");
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));
        Map<String, Object> resources = new LinkedHashMap<>();
        resources.put("plugin1", "data1");
        resources.put("plugin2", 42);
        src.setResources(resources);

        IWorkbookData back = roundtrip(src);

        assertThat(back.getResources()).isNotNull();
        assertThat(back.getResources()).isInstanceOf(Map.class);
        @SuppressWarnings("unchecked")
        Map<String, Object> backRes = (Map<String, Object>) back.getResources();
        assertThat(backRes).containsKeys("plugin1", "plugin2");
        assertThat(backRes.get("plugin1")).isEqualTo("data1");
        assertThat(((Number) backRes.get("plugin2")).intValue()).isEqualTo(42);
    }

    // ============================================================
    // 8. cell custom field
    // ============================================================

    @Test
    void should_roundtrip_cell_custom_field() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<String, Object> custom = new LinkedHashMap<>();
        custom.put("key", "value");
        custom.put("num", 100);
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("x").setT(CellValueType.STRING).setCustom(custom));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        ICellData c = back.getSheets().get("s1").getCellData().get(0).get(0);
        assertThat(c.getCustom()).isNotNull();
        assertThat(c.getCustom()).containsKeys("key", "num");
        assertThat(c.getCustom().get("key")).isEqualTo("value");
        assertThat(((Number) c.getCustom().get("num")).intValue()).isEqualTo(100);
    }

    // ============================================================
    // 9. external xlsx without sidecar
    // ============================================================

    @Test
    void should_read_external_xlsx_without_sidecar_gracefully() throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("ExternalSheet");
            sheet.createRow(0).createCell(0).setCellValue("hello");
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            bytes = out.toByteArray();
        }

        IWorkbookData data = UniverXlsx.read(new ByteArrayInputStream(bytes));

        assertThat(data).isNotNull();
        assertThat(data.getSheets()).containsKey("ExternalSheet");
        ICellData a1 = data.getSheets().get("ExternalSheet").getCellData().get(0).get(0);
        assertThat(a1).isNotNull();
        assertThat(a1.getV()).isEqualTo("hello");
        assertThat(data.getLocale()).isEqualTo("enUS");
    }

    // ============================================================
    // 9b. external xlsx populates styles map
    // ============================================================

    @Test
    void should_populate_styles_map_when_reading_external_xlsx() throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Styled");
            org.apache.poi.xssf.usermodel.XSSFCellStyle s1 = wb.createCellStyle();
            org.apache.poi.xssf.usermodel.XSSFFont f1 = wb.createFont();
            f1.setFontName("Arial");
            f1.setFontHeightInPoints((short) 14);
            s1.setFont(f1);

            org.apache.poi.xssf.usermodel.XSSFCellStyle s2 = wb.createCellStyle();
            org.apache.poi.xssf.usermodel.XSSFFont f2 = wb.createFont();
            f2.setFontName("Courier New");
            f2.setFontHeightInPoints((short) 10);
            f2.setBold(true);
            s2.setFont(f2);

            org.apache.poi.xssf.usermodel.XSSFRow row = sheet.createRow(0);
            org.apache.poi.xssf.usermodel.XSSFCell c1 = row.createCell(0);
            c1.setCellValue("x");
            c1.setCellStyle(s1);
            org.apache.poi.xssf.usermodel.XSSFCell c2 = row.createCell(1);
            c2.setCellValue("y");
            c2.setCellStyle(s2);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            bytes = out.toByteArray();
        }

        IWorkbookData data = UniverXlsx.read(new ByteArrayInputStream(bytes));

        assertThat(data.getStyles()).isNotNull();
        assertThat(data.getStyles()).isNotEmpty();
        IWorksheetData ws = data.getSheets().get("Styled");
        ICellData a1 = ws.getCellData().get(0).get(0);
        ICellData b1 = ws.getCellData().get(0).get(1);
        assertThat(a1.getS()).isInstanceOf(String.class);
        assertThat(b1.getS()).isInstanceOf(String.class);
        assertThat(data.getStyles()).containsKey((String) a1.getS());
        assertThat(data.getStyles()).containsKey((String) b1.getS());
    }

    // ============================================================
    // 10. empty workbook with one empty sheet
    // ============================================================

    @Test
    void should_roundtrip_empty_workbook_with_one_empty_sheet() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("EmptySheet");
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        assertThat(back.getSheets()).containsKey("s1");
        assertThat(back.getSheets().get("s1").getName()).isEqualTo("EmptySheet");
    }

    // ============================================================
    // 11. sidecar disabled
    // ============================================================

    @Test
    void should_roundtrip_with_sidecar_disabled() throws Exception {
        UniverXlsxOptions opts = UniverXlsxOptions.builder().writeSidecar(false).build();
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("NoSidecar");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("hello").setT(CellValueType.STRING));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        byte[] bytes = writeBytes(src, opts);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(bytes), opts);

        assertThat(back.getSheets()).containsKey("NoSidecar");
        ICellData a1 = back.getSheets().get("NoSidecar").getCellData().get(0).get(0);
        assertThat(a1).isNotNull();
        assertThat(a1.getV()).isEqualTo("hello");
    }

    // ============================================================
    // 12. force-text cell
    // ============================================================

    @Test
    void should_roundtrip_force_text_cell() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData().setV("012.0").setT(CellValueType.FORCE_TEXT));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        IWorkbookData back = roundtrip(src);

        ICellData c = back.getSheets().get("s1").getCellData().get(0).get(0);
        assertThat(c.getT()).isEqualTo(CellValueType.FORCE_TEXT);
        assertThat(c.getV()).isEqualTo("012.0");
    }
}
