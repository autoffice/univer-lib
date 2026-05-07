package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.IColorStyle;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextDecoration;
import io.github.autoffice.univer.util.ColorUtils;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * 富文本转换器。
 * Rich text converter between IDocumentData and XSSFRichTextString.
 */
public final class RichTextConverter {

    /** 承载 POI 工作簿的引用，用于创建字体 / workbook reference for font creation. */
    private final XSSFWorkbook wb;

    public RichTextConverter(XSSFWorkbook wb) {
        this.wb = wb;
    }

    // ============================================================
    // 公共 API / Public API
    // ============================================================

    /**
     * 将 IDocumentData 转为 POI 富文本。
     * Convert IDocumentData to POI rich text string.
     */
    public XSSFRichTextString toPoi(IDocumentData p) {
        if (p == null || p.getBody() == null) {
            return new XSSFRichTextString("");
        }
        String dataStream = p.getBody().getDataStream();
        if (dataStream == null || dataStream.isEmpty()) {
            return new XSSFRichTextString("");
        }
        XSSFRichTextString rts = new XSSFRichTextString(dataStream);
        List<IDocumentData.TextRun> runs = p.getBody().getTextRuns();
        if (runs != null) {
            for (IDocumentData.TextRun run : runs) {
                if (run == null || run.getSt() == null || run.getEd() == null) {
                    continue;
                }
                XSSFFont font = createFont(run.getTs() == null ? new IStyleData() : run.getTs());
                rts.applyFont(run.getSt(), run.getEd(), font);
            }
        }
        return rts;
    }

    /**
     * 将 POI 富文本转为 IDocumentData。
     * Convert POI rich text string to IDocumentData.
     */
    public IDocumentData fromPoi(XSSFRichTextString rts) {
        String dataStream = rts == null || rts.getString() == null ? "" : rts.getString();

        List<IDocumentData.TextRun> runs = new ArrayList<>();
        if (rts != null) {
            int count = rts.numFormattingRuns();
            for (int i = 0; i < count; i++) {
                int st = rts.getIndexOfFormattingRun(i);
                int len = rts.getLengthOfFormattingRun(i);
                int ed = st + len;
                XSSFFont font = rts.getFontOfFormattingRun(i);
                if (font == null) {
                    continue;
                }
                IStyleData ts = fontToStyle(font);
                runs.add(new IDocumentData.TextRun().setSt(st).setEd(ed).setTs(ts));
            }
        }

        List<IDocumentData.Paragraph> paragraphs = new ArrayList<>();
        int idx = dataStream.indexOf('\n');
        while (idx >= 0) {
            paragraphs.add(new IDocumentData.Paragraph().setStartIndex(idx));
            idx = dataStream.indexOf('\n', idx + 1);
        }

        return new IDocumentData().setBody(new IDocumentData.Body()
            .setDataStream(dataStream)
            .setTextRuns(runs)
            .setParagraphs(paragraphs));
    }

    // ============================================================
    // 字体辅助 / Font helpers
    // ============================================================

    /**
     * 根据 IStyleData 中与字体相关的子集创建 XSSFFont。
     * Create an XSSFFont from the font-related subset of IStyleData.
     */
    private XSSFFont createFont(IStyleData ts) {
        XSSFFont font = wb.createFont();
        if (ts.getFf() != null) {
            font.setFontName(ts.getFf());
        }
        if (ts.getFs() != null) {
            font.setFontHeightInPoints(ts.getFs().shortValue());
        }
        if (ts.getBl() == BooleanNumber.TRUE) {
            font.setBold(true);
        }
        if (ts.getIt() == BooleanNumber.TRUE) {
            font.setItalic(true);
        }
        if (ts.getUl() != null && ts.getUl().getS() == BooleanNumber.TRUE) {
            font.setUnderline(FontUnderline.SINGLE.getByteValue());
        }
        if (ts.getSt() != null && ts.getSt().getS() == BooleanNumber.TRUE) {
            font.setStrikeout(true);
        }
        if (ts.getCl() != null && ts.getCl().getRgb() != null) {
            font.setColor(new XSSFColor(ColorUtils.rgbHexToArgb(ts.getCl().getRgb()), null));
        }
        return font;
    }

    /**
     * 将 XSSFFont 读回仅包含字体字段的 IStyleData。
     * Read XSSFFont back to an IStyleData holding only font-related fields.
     */
    private IStyleData fontToStyle(XSSFFont font) {
        IStyleData ts = new IStyleData();
        if (font.getFontName() != null) {
            ts.setFf(font.getFontName());
        }
        if (font.getFontHeightInPoints() > 0) {
            ts.setFs((int) font.getFontHeightInPoints());
        }
        if (font.getBold()) {
            ts.setBl(BooleanNumber.TRUE);
        }
        if (font.getItalic()) {
            ts.setIt(BooleanNumber.TRUE);
        }
        if (font.getUnderline() != (byte) 0) {
            ts.setUl(new ITextDecoration().setS(BooleanNumber.TRUE));
        }
        if (font.getStrikeout()) {
            ts.setSt(new ITextDecoration().setS(BooleanNumber.TRUE));
        }
        XSSFColor color = font.getXSSFColor();
        if (color != null) {
            byte[] argb = color.getARGB();
            if (argb != null && argb.length >= 4) {
                String hex = ColorUtils.argbToRgbHex(argb);
                if (hex != null) {
                    ts.setCl(new IColorStyle().setRgb(hex));
                }
            }
        }
        return ts;
    }
}
