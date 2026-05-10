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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * 富文本转换器。
 * Rich text converter between IDocumentData and XSSFRichTextString.
 * <p>
 * 超链接通过 {@code IDocumentData.Body.customRanges} 表达；模型未定义该字段，
 * 这里通过 {@link io.github.autoffice.univer.model.AbstractUniverModel#putExtra} 写入 extras，
 * Jackson 会把它平铺回 body，与 Univer TS 侧结构一致。
 * Hyperlinks are expressed via {@code IDocumentData.Body.customRanges}. The POJO does not declare
 * the field; this converter writes it through {@code extras} so Jackson emits it at the body level
 * (Univer TS shape).
 */
public final class RichTextConverter {

    /** body.customRanges 字段名 / field name for customRanges on body. */
    public static final String CUSTOM_RANGES_KEY = "customRanges";
    /** rangeType = 0 表示超链接 / rangeType 0 = hyperlink. */
    public static final int RANGE_TYPE_HYPERLINK = 0;

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

    // ============================================================
    // 超链接辅助 / Hyperlink helpers
    // ============================================================

    /**
     * 读取 {@code body.customRanges} 列表；若不存在则返回空列表。
     * Read {@code body.customRanges} from extras; return empty list if absent.
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, Object>> customRangesOf(IDocumentData p) {
        if (p == null || p.getBody() == null) {
            return new ArrayList<>();
        }
        Object raw = p.getBody().getExtras().get(CUSTOM_RANGES_KEY);
        if (raw instanceof List) {
            return (List<Map<String, Object>>) raw;
        }
        return new ArrayList<>();
    }

    /**
     * 返回文档中第一个超链接 custom range 的 url（没有则返回 null）。
     * Return the url of the first hyperlink custom range, or null if none exists.
     */
    public static String firstHyperlinkUrl(IDocumentData p) {
        for (Map<String, Object> range : customRangesOf(p)) {
            Object rt = range.get("rangeType");
            if (rt instanceof Number && ((Number) rt).intValue() != RANGE_TYPE_HYPERLINK) {
                continue;
            }
            Object props = range.get("properties");
            if (props instanceof Map) {
                Object url = ((Map<String, Object>) props).get("url");
                if (url instanceof String && !((String) url).isEmpty()) {
                    return (String) url;
                }
            }
        }
        return null;
    }

    /**
     * 在 body.customRanges 上追加一段超链接 range；若列表不存在则创建。
     * 起止索引按照 Univer dataStream 约定（"\r\n" 段落终止符），范围闭区间。
     * Append a hyperlink custom range to body.customRanges (create list if absent). Indices follow
     * Univer dataStream convention (end-exclusive style used in showcase payloads).
     */
    @SuppressWarnings("unchecked")
    public static void attachHyperlink(IDocumentData p, String url, int startIndex, int endIndex) {
        if (p == null || p.getBody() == null || url == null || url.isEmpty()) {
            return;
        }
        Map<String, Object> extras = p.getBody().getExtras();
        Object raw = extras.get(CUSTOM_RANGES_KEY);
        List<Map<String, Object>> ranges;
        if (raw instanceof List) {
            ranges = (List<Map<String, Object>>) raw;
        } else {
            ranges = new ArrayList<>();
            p.getBody().putExtra(CUSTOM_RANGES_KEY, ranges);
        }
        String rangeId = UUID.randomUUID().toString();
        Map<String, Object> range = new LinkedHashMap<>();
        range.put("rangeId", rangeId);
        range.put("rangeType", RANGE_TYPE_HYPERLINK);
        range.put("startIndex", startIndex);
        range.put("endIndex", endIndex);
        Map<String, Object> props = new LinkedHashMap<>();
        props.put("url", url);
        props.put("refId", rangeId);
        range.put("properties", props);
        ranges.add(range);
    }
}
