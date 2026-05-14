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

import io.github.autoffice.univer.UniverXlsx;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
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

class HyperlinkConverterTest {

    private WorksheetConverter newConverter(XSSFWorkbook wb,
                                            Map<String, String> sheetIdToName,
                                            Map<String, String> sheetNameToId) {
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        WorksheetConverter wsc = new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());
        wsc.setSheetIdMaps(sheetIdToName, sheetNameToId);
        return wsc;
    }

    private IDocumentData richTextWithLink(String text, String url) {
        IDocumentData doc = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream(text + "\r\n")
                .setTextRuns(Arrays.asList(
                        new IDocumentData.TextRun().setSt(0).setEd(text.length())))
                .setParagraphs(Arrays.asList(
                        new IDocumentData.Paragraph().setStartIndex(text.length()))));
        RichTextConverter.attachHyperlink(doc, url, 0, text.length());
        return doc;
    }

    @Test
    void should_write_external_url_hyperlink_to_xlsx_cell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb, Collections.emptyMap(), Collections.emptyMap());
            XSSFSheet sheet = wb.createSheet("S1");

            IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData()
                    .setV("Univer")
                    .setT(CellValueType.STRING)
                    .setP(richTextWithLink("Univer", "https://univer.ai")));
            ws.getCellData().put(0, row);

            wsc.writeSheet(sheet, ws);

            XSSFCell cell = sheet.getRow(0).getCell(0);
            assertThat(cell.getStringCellValue()).isEqualTo("Univer");
            XSSFHyperlink h = cell.getHyperlink();
            assertThat(h).isNotNull();
            assertThat(h.getType()).isEqualTo(HyperlinkType.URL);
            assertThat(h.getAddress()).isEqualTo("https://univer.ai");
        }
    }

    @Test
    void should_translate_internal_gid_url_to_quoted_sheet_ref_on_write() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // 创建被引用的目标 sheet（名字含空格，验证引号转义）/ Target sheet name has space → must be quoted.
            wb.createSheet("Other Sheet");

            Map<String, String> idToName = new LinkedHashMap<>();
            idToName.put("target-sid", "Other Sheet");
            WorksheetConverter wsc = newConverter(wb, idToName, Collections.emptyMap());
            XSSFSheet sheet = wb.createSheet("Home");

            IWorksheetData ws = new IWorksheetData().setId("home").setName("Home");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData()
                    .setV("go")
                    .setT(CellValueType.STRING)
                    .setP(richTextWithLink("go", "#gid=target-sid&range=B2")));
            ws.getCellData().put(0, row);

            wsc.writeSheet(sheet, ws);

            XSSFHyperlink h = sheet.getRow(0).getCell(0).getHyperlink();
            assertThat(h).isNotNull();
            assertThat(h.getType()).isEqualTo(HyperlinkType.DOCUMENT);
            assertThat(h.getAddress()).isEqualTo("'Other Sheet'!B2");
        }
    }

    @Test
    void should_read_xlsx_url_hyperlink_into_custom_range() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("S1");
            XSSFCell cell = sheet.createRow(0).createCell(0);
            cell.setCellValue("docs");
            XSSFHyperlink h = (XSSFHyperlink) wb.getCreationHelper().createHyperlink(HyperlinkType.URL);
            h.setAddress("https://docs.univer.ai");
            cell.setHyperlink(h);

            WorksheetConverter wsc = newConverter(wb, Collections.emptyMap(), Collections.emptyMap());
            IWorksheetData back = wsc.readSheet(sheet);

            ICellData data = back.getCellData().get(0).get(0);
            assertThat(data.getP()).as("hyperlink should populate p").isNotNull();
            String url = RichTextConverter.firstHyperlinkUrl(data.getP());
            assertThat(url).isEqualTo("https://docs.univer.ai");
        }
    }

    @Test
    void should_read_internal_xlsx_hyperlink_back_to_gid_url() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Other");
            XSSFSheet home = wb.createSheet("Home");
            XSSFCell cell = home.createRow(0).createCell(0);
            cell.setCellValue("go");
            XSSFHyperlink h = (XSSFHyperlink) wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
            h.setAddress("Other!B2");
            cell.setHyperlink(h);

            Map<String, String> nameToId = new LinkedHashMap<>();
            nameToId.put("Other", "target-sid");
            nameToId.put("Home", "home");
            WorksheetConverter wsc = newConverter(wb, Collections.emptyMap(), nameToId);
            IWorksheetData back = wsc.readSheet(home);

            ICellData data = back.getCellData().get(0).get(0);
            String url = RichTextConverter.firstHyperlinkUrl(data.getP());
            assertThat(url).isEqualTo("#gid=target-sid&range=B2");
        }
    }

    @Test
    void should_roundtrip_external_hyperlink_through_xlsx_only() throws Exception {
        // 关闭 sidecar，确保 hyperlink 是从 xlsx 原生载体而非边车 JSON 还原。
        // Disable sidecar to prove the hyperlink survives via xlsx, not the JSON sidecar.
        UniverXlsxOptions opts = UniverXlsxOptions.builder().writeSidecar(false).build();

        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData()
                .setV("Univer")
                .setT(CellValueType.STRING)
                .setP(richTextWithLink("Univer", "https://univer.ai")));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf, opts);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));

        // 任意一个 sheet 即可（外部 xlsx 读取时 sid = sheet name）
        IWorksheetData backWs = back.getSheets().values().iterator().next();
        ICellData backCell = backWs.getCellData().get(0).get(0);
        assertThat(RichTextConverter.firstHyperlinkUrl(backCell.getP()))
                .isEqualTo("https://univer.ai");
    }

    @Test
    void should_roundtrip_internal_gid_hyperlink_through_xlsx_only() throws Exception {
        UniverXlsxOptions opts = UniverXlsxOptions.builder().writeSidecar(false).build();

        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData home = new IWorksheetData().setId("home").setName("Home");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData()
                .setV("go")
                .setT(CellValueType.STRING)
                .setP(richTextWithLink("go", "#gid=target-sid&range=B2")));
        home.getCellData().put(0, row);
        IWorksheetData target = new IWorksheetData().setId("target-sid").setName("Other");
        src.getSheets().put("home", home);
        src.getSheets().put("target-sid", target);
        src.setSheetOrder(Arrays.asList("home", "target-sid"));

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf, opts);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));

        // 无 sidecar 时 sid 退化为 sheet name；hyperlink 也据此还原。
        // Without sidecar sid degrades to sheet name; the link follows suit.
        IWorksheetData backHome = back.getSheets().get("Home");
        assertThat(backHome).isNotNull();
        ICellData backCell = backHome.getCellData().get(0).get(0);
        String url = RichTextConverter.firstHyperlinkUrl(backCell.getP());
        assertThat(url).isEqualTo("#gid=Other&range=B2");
    }

    @Test
    void should_preserve_sidecar_customRanges_through_roundtrip() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("wb").setAppVersion("0.10.2");
        IWorksheetData ws = new IWorksheetData().setId("s1").setName("S1");
        Map<Integer, ICellData> row = new LinkedHashMap<>();
        row.put(0, new ICellData()
                .setV("Univer")
                .setT(CellValueType.STRING)
                .setP(richTextWithLink("Univer", "https://univer.ai")));
        ws.getCellData().put(0, row);
        src.getSheets().put("s1", ws);
        src.setSheetOrder(Collections.singletonList("s1"));

        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        UniverXlsx.write(src, buf);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(buf.toByteArray()));

        ICellData backCell = back.getSheets().get("s1").getCellData().get(0).get(0);
        List<Map<String, Object>> ranges = RichTextConverter.customRangesOf(backCell.getP());
        assertThat(ranges).hasSize(1);
        Map<String, Object> r0 = ranges.get(0);
        assertThat(r0).containsEntry("rangeType", 0);
        @SuppressWarnings("unchecked")
        Map<String, Object> props = (Map<String, Object>) r0.get("properties");
        assertThat(props).containsEntry("url", "https://univer.ai");
        // refId 与 rangeId 一致（与 Univer 约定）
        assertThat(props.get("refId")).isEqualTo(r0.get("rangeId"));
    }
}
