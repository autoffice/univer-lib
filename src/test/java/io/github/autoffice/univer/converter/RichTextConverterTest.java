package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.util.Arrays;

import static org.assertj.core.api.Assertions.assertThat;

class RichTextConverterTest {

    @Test
    void should_roundtrip_runs_and_paragraphs() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData.TextRun run1 = new IDocumentData.TextRun().setSt(0).setEd(5)
                .setTs(new IStyleData().setFf("Arial").setFs(10));
            IDocumentData.TextRun run2 = new IDocumentData.TextRun().setSt(6).setEd(11)
                .setTs(new IStyleData().setFf("Arial").setFs(14)
                    .setCl(new IColorStyle().setRgb("#ff0000")));
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("hello\nworld")
                .setTextRuns(Arrays.asList(run1, run2))
                .setParagraphs(Arrays.asList(new IDocumentData.Paragraph().setStartIndex(5))));

            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts.getString()).isEqualTo("hello\nworld");

            IDocumentData back = rc.fromPoi(rts);
            assertThat(back.getBody().getDataStream()).isEqualTo("hello\nworld");
            assertThat(back.getBody().getTextRuns()).hasSizeGreaterThanOrEqualTo(2);
            assertThat(back.getBody().getParagraphs()).hasSize(1);
            assertThat(back.getBody().getParagraphs().get(0).getStartIndex()).isEqualTo(5);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_plain_text_without_runs() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("plain text"));
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts.getString()).isEqualTo("plain text");
            IDocumentData back = rc.fromPoi(rts);
            assertThat(back.getBody().getDataStream()).isEqualTo("plain text");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_preserve_bold_italic_in_runs() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData.TextRun run = new IDocumentData.TextRun().setSt(0).setEd(4)
                .setTs(new IStyleData().setBl(BooleanNumber.TRUE).setIt(BooleanNumber.TRUE));
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("bold")
                .setTextRuns(Arrays.asList(run)));
            XSSFRichTextString rts = rc.toPoi(p);
            IDocumentData back = rc.fromPoi(rts);
            assertThat(back.getBody().getTextRuns()).isNotEmpty();
            IStyleData ts = back.getBody().getTextRuns().get(0).getTs();
            assertThat(ts.getBl()).isEqualTo(BooleanNumber.TRUE);
            assertThat(ts.getIt()).isEqualTo(BooleanNumber.TRUE);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_null_body_gracefully() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            XSSFRichTextString rts = rc.toPoi(new IDocumentData());
            assertThat(rts.getString()).isEmpty();
            IDocumentData back = rc.fromPoi(new XSSFRichTextString(""));
            assertThat(back.getBody().getDataStream()).isEmpty();
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
