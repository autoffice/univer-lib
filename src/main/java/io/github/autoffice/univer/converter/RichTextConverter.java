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
     * <p>Univer 的 dataStream 以 "\r\n" 作为段落终止符，写回 xlsx 时需剥离并切换为换行。
     */
    public XSSFRichTextString toPoi(IDocumentData p) {
        if (p == null || p.getBody() == null) {
            return new XSSFRichTextString("");
        }
        String dataStream = p.getBody().getDataStream();
        if (dataStream == null || dataStream.isEmpty()) {
            return new XSSFRichTextString("");
        }
        // 去掉 Univer 段落终止符 "\r\n"，把内部保留的 "\r\n" 合并成 "\n"
        String plain = dataStream.replace("\r\n", "\n");
        while (plain.endsWith("\n")) {
            plain = plain.substring(0, plain.length() - 1);
        }
        XSSFRichTextString rts = new XSSFRichTextString(plain);
        List<IDocumentData.TextRun> runs = p.getBody().getTextRuns();
        if (runs != null) {
            int plainLen = plain.length();
            for (IDocumentData.TextRun run : runs) {
                if (run == null || run.getSt() == null || run.getEd() == null) {
                    continue;
                }
                int st = Math.max(0, Math.min(run.getSt(), plainLen));
                int ed = Math.max(st, Math.min(run.getEd(), plainLen));
                if (st >= ed) {
                    continue;
                }
                XSSFFont font = createFont(run.getTs() == null ? new IStyleData() : run.getTs());
                rts.applyFont(st, ed, font);
            }
        }
        return rts;
    }

    /**
     * 将 POI 富文本转为 IDocumentData，格式严格匹配 Univer Docs 约定。
     * Convert POI rich text to IDocumentData, strictly matching Univer's doc-stream convention:
     * <ul>
     *   <li>dataStream 以 "\r\n" 作为段落终止符；每段结尾追加 "\r\n"。</li>
     *   <li>paragraphs 至少一项，startIndex 指向段落尾字符（即 "\r" 前的位置）。</li>
     *   <li>textRuns 的 st/ed 基于 dataStream 的索引。</li>
     * </ul>
     */
    public IDocumentData fromPoi(XSSFRichTextString rts) {
        String raw = rts == null || rts.getString() == null ? "" : rts.getString();

        // 按 \n 切段（xlsx 中段内换行也是 \n），每段后追加 "\r\n"
        StringBuilder streamBuilder = new StringBuilder();
        List<IDocumentData.Paragraph> paragraphs = new ArrayList<>();
        String[] lines = raw.split("\n", -1);
        int cursor = 0;
        for (String line : lines) {
            streamBuilder.append(line);
            cursor += line.length();
            paragraphs.add(new IDocumentData.Paragraph().setStartIndex(cursor));
            streamBuilder.append("\r\n");
            cursor += 2;
        }
        String dataStream = streamBuilder.toString();

        // 计算 raw 索引 → dataStream 索引的偏移（每经过一个 "\n" 就 +1，因为 "\n" 变成了 "\r\n"）
        List<IDocumentData.TextRun> runs = new ArrayList<>();
        if (rts != null) {
            int count = rts.numFormattingRuns();
            for (int i = 0; i < count; i++) {
                int rawSt = rts.getIndexOfFormattingRun(i);
                int rawLen = rts.getLengthOfFormattingRun(i);
                int rawEd = rawSt + rawLen;
                XSSFFont font = rts.getFontOfFormattingRun(i);
                if (font == null) {
                    continue;
                }
                IStyleData ts = fontToStyle(font);
                int st = mapRawIndex(raw, rawSt);
                int ed = mapRawIndex(raw, rawEd);
                runs.add(new IDocumentData.TextRun().setSt(st).setEd(ed).setTs(ts));
            }
        }

        return new IDocumentData().setBody(new IDocumentData.Body()
            .setDataStream(dataStream)
            .setTextRuns(runs)
            .setParagraphs(paragraphs));
    }

    /** 把 raw 索引（仅含 "\n"）换算为 dataStream 索引（"\n" 已变成 "\r\n"）。 */
    private static int mapRawIndex(String raw, int rawIdx) {
        int extra = 0;
        int max = Math.min(rawIdx, raw.length());
        for (int i = 0; i < max; i++) {
            if (raw.charAt(i) == '\n') {
                extra++;
            }
        }
        return rawIdx + extra;
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
            byte[] argb = ColorUtils.rgbHexToArgb(ts.getCl().getRgb());
            if (argb != null) {
                font.setColor(new XSSFColor(argb, null));
            }
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
