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

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.CellValueType;
import io.github.autoffice.univer.model.ICellData;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IRowData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * WorksheetConverter 边角分支：null src、zoomRatio、hyperlink-only doc、行隐藏 hd、不同 hyperlink 类型。
 */
class WorksheetConverterBranchTest {

    private WorksheetConverter newConverter(XSSFWorkbook wb) {
        StyleConverter sc = new StyleConverter(wb);
        CellConverter cc = new CellConverter(sc);
        RichTextConverter rc = new RichTextConverter(wb);
        SharedFormulaRegistry sfr = new SharedFormulaRegistry();
        return new WorksheetConverter(wb, sc, cc, rc, sfr, UniverXlsxOptions.defaults());
    }

    @Test
    void should_skip_when_src_is_null() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            wsc.writeSheet(sh, null);
            // 不抛错；sheet 仍存在
            assertThat(wb.getNumberOfSheets()).isEqualTo(1);
        }
    }

    @Test
    void should_apply_zoom_ratio() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S")
                    .setZoomRatio(1.5);
            wsc.writeSheet(sh, ws);
            // POI getZoom() 没有公开 API，但 setZoom 不会抛错
            assertThat(sh.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale())
                    .isEqualTo(150);
        }
    }

    @Test
    void should_skip_zoom_ratio_when_zero() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S")
                    .setZoomRatio(0.0);
            wsc.writeSheet(sh, ws);
            // 不应抛错
            assertThat(sh.getSheetName()).isEqualTo("S");
        }
    }

    @Test
    void should_handle_hyperlink_only_doc_no_text() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            // dataStream 仅是 "\r\n"，表示无文本仅承载 hyperlink
            IDocumentData doc = new IDocumentData().setBody(
                    new IDocumentData.Body().setDataStream("\r\n"));
            RichTextConverter.attachHyperlink(doc, "https://example.com", 0, 0);
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setP(doc));
            ws.getCellData().put(0, row);
            wsc.writeSheet(sh, ws);
            XSSFCell c = sh.getRow(0).getCell(0);
            // BLANK 单元格但有 hyperlink
            assertThat(c.getHyperlink()).isNotNull();
            assertThat(c.getHyperlink().getAddress()).isEqualTo("https://example.com");
        }
    }

    @Test
    void should_handle_email_hyperlink() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IDocumentData doc = richTextWithLink("contact", "mailto:hello@example.com");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("contact").setT(CellValueType.STRING).setP(doc));
            ws.getCellData().put(0, row);
            wsc.writeSheet(sh, ws);
            XSSFCell c = sh.getRow(0).getCell(0);
            assertThat(c.getHyperlink()).isNotNull();
            assertThat(c.getHyperlink().getType()).isEqualTo(HyperlinkType.EMAIL);
        }
    }

    @Test
    void should_handle_file_hyperlink() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IDocumentData doc = richTextWithLink("doc", "file:///tmp/x.txt");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("doc").setT(CellValueType.STRING).setP(doc));
            ws.getCellData().put(0, row);
            wsc.writeSheet(sh, ws);
            XSSFCell c = sh.getRow(0).getCell(0);
            assertThat(c.getHyperlink()).isNotNull();
            assertThat(c.getHyperlink().getType()).isEqualTo(HyperlinkType.FILE);
        }
    }

    @Test
    void should_handle_email_via_at_symbol_without_mailto() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IDocumentData doc = richTextWithLink("c", "name@example.com");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("c").setT(CellValueType.STRING).setP(doc));
            ws.getCellData().put(0, row);
            wsc.writeSheet(sh, ws);
            XSSFCell c = sh.getRow(0).getCell(0);
            assertThat(c.getHyperlink().getType()).isEqualTo(HyperlinkType.EMAIL);
        }
    }

    @Test
    void should_apply_row_hidden_via_hd_field() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, IRowData> rowData = new LinkedHashMap<>();
            rowData.put(0, new IRowData().setHd(BooleanNumber.TRUE));
            ws.setRowData(rowData);
            wsc.writeSheet(sh, ws);
            assertThat(sh.getRow(0).getZeroHeight()).isTrue();
        }
    }

    @Test
    void should_skip_hyperlink_with_invalid_address() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            WorksheetConverter wsc = newConverter(wb);
            XSSFSheet sh = wb.createSheet("S");
            // 内部 link 但目标 sheet 不存在 → translateUrlForXlsx 返回非空但 createHyperlink 可能拒绝
            IDocumentData doc = richTextWithLink("x", "#gid=ghost-sid&range=A1");
            IWorksheetData ws = new IWorksheetData().setId("s").setName("S");
            Map<Integer, ICellData> row = new LinkedHashMap<>();
            row.put(0, new ICellData().setV("x").setT(CellValueType.STRING).setP(doc));
            ws.getCellData().put(0, row);
            wsc.setSheetIdMaps(Collections.emptyMap(), Collections.emptyMap());
            wsc.writeSheet(sh, ws);
            // 不抛错就行
            assertThat(sh.getRow(0).getCell(0)).isNotNull();
        }
    }

    @Test
    void should_handle_xlsx_hyperlink_with_location_only() throws Exception {
        // POI hyperlink 的 address 为 null 但 location 不为 null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            XSSFCell cell = sh.createRow(0).createCell(0);
            cell.setCellValue("anchor");
            XSSFHyperlink h = (XSSFHyperlink) wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
            h.setLocation("Other!A1");
            cell.setHyperlink(h);
            WorksheetConverter wsc = newConverter(wb);
            IWorksheetData back = wsc.readSheet(sh);
            ICellData data = back.getCellData().get(0).get(0);
            assertThat(data.getP()).isNotNull();
            String url = RichTextConverter.firstHyperlinkUrl(data.getP());
            assertThat(url).isNotNull();
        }
    }

    @Test
    void should_skip_xlsx_hyperlink_when_address_blank() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            XSSFCell cell = sh.createRow(0).createCell(0);
            cell.setCellValue("anchor");
            XSSFHyperlink h = (XSSFHyperlink) wb.getCreationHelper().createHyperlink(HyperlinkType.URL);
            h.setAddress("");
            cell.setHyperlink(h);
            WorksheetConverter wsc = newConverter(wb);
            IWorksheetData back = wsc.readSheet(sh);
            ICellData data = back.getCellData().get(0).get(0);
            // hyperlink 地址为空 → 不应在 p 上附加 url
            String url = data.getP() == null ? null : RichTextConverter.firstHyperlinkUrl(data.getP());
            assertThat(url).isNull();
        }
    }

    private IDocumentData richTextWithLink(String text, String url) {
        IDocumentData doc = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream(text + "\r\n")
                .setParagraphs(Collections.singletonList(
                        new IDocumentData.Paragraph().setStartIndex(text.length()))));
        RichTextConverter.attachHyperlink(doc, url, 0, text.length());
        return doc;
    }
}
