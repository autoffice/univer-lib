package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.IBorderData;
import io.github.autoffice.univer.model.IBorderStyleData;
import io.github.autoffice.univer.model.IColorStyle;
import io.github.autoffice.univer.model.INumfmtLocal;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextDecoration;
import io.github.autoffice.univer.model.ITextRotation;
import io.github.autoffice.univer.util.ColorUtils;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 样式转换器：在 Univer IStyleData 与 POI XSSFCellStyle/XSSFFont 之间双向映射。
 * Bidirectional converter between Univer IStyleData and POI XSSFCellStyle/XSSFFont.
 */
public final class StyleConverter {

    /** 承载 POI 工作簿的引用 / the target workbook. */
    private final XSSFWorkbook wb;
    /** 写路径样式缓存，键为 styleIdOf / write-path cache keyed by styleIdOf. */
    private final Map<String, XSSFCellStyle> cache = new HashMap<>();
    /** quotePrefix 变体缓存，键为 styleIdOf+"#qp" / cache for quote-prefix style variants. */
    private final Map<String, XSSFCellStyle> quotePrefixCache = new HashMap<>();
    /** 读路径样式去重注册表：styleIdOf → IStyleData / registry of styles observed via styleIdOf. */
    private final Map<String, IStyleData> idToStyle = new LinkedHashMap<>();

    public StyleConverter(XSSFWorkbook wb) {
        this.wb = wb;
    }

    // ============================================================
    // 公共 API / Public API
    // ============================================================

    /**
     * IStyleData → XSSFCellStyle（写路径，结果缓存）。
     * Convert IStyleData to XSSFCellStyle (cached by styleIdOf).
     */
    public XSSFCellStyle toPoiStyle(IStyleData s) {
        IStyleData src = s == null ? new IStyleData() : s;
        String id = styleIdOf(src);
        XSSFCellStyle cached = cache.get(id);
        if (cached != null) {
            return cached;
        }
        XSSFFont font = wb.createFont();
        applyFont(src, font);
        XSSFCellStyle style = wb.createCellStyle();
        applyBackground(src, style);
        applyBorders(src, style);
        applyAlignment(src, style);
        applyWrap(src, style);
        applyRotation(src, style);
        applyNumFmt(src, style);
        style.setFont(font);
        cache.put(id, style);
        return style;
    }

    /**
     * XSSFCellStyle → IStyleData（读路径，默认值被跳过）。
     * Convert XSSFCellStyle back to IStyleData (defaults are skipped).
     */
    public IStyleData fromPoiStyle(XSSFCellStyle cs) {
        IStyleData out = new IStyleData();
        if (cs == null || cs.getIndex() == 0) {
            return out;
        }
        readFont(cs, out);
        readBackground(cs, out);
        readBorders(cs, out);
        readAlignment(cs, out);
        readWrap(cs, out);
        readRotation(cs, out);
        readNumFmt(cs, out);
        return out;
    }

    /**
     * 为 IStyleData 生成稳定的 16 位 16 进制哈希（SHA-256 前缀），用于去重。
     * Generate a stable 16-hex-char id for IStyleData (SHA-256 prefix).
     */
    public String styleIdOf(IStyleData s) {
        IStyleData src = s == null ? new IStyleData() : s;
        try {
            String json = JsonMapper.get().writeValueAsString(src);
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] hash = md.digest(json.getBytes(StandardCharsets.UTF_8));
            StringBuilder sb = new StringBuilder(16);
            for (int i = 0; i < 8; i++) {
                sb.append(String.format("%02x", hash[i] & 0xFF));
            }
            String id = sb.toString();
            // 登记到读路径注册表，便于 WorkbookConverter 回填 IWorkbookData.styles
            // Register into read-path registry so WorkbookConverter can populate IWorkbookData.styles.
            idToStyle.putIfAbsent(id, src);
            return id;
        } catch (Exception e) {
            throw new IllegalStateException("Failed to hash style: " + e.getMessage(), e);
        }
    }

    /**
     * 返回读路径累计的 styleId → IStyleData 注册表（插入顺序）。
     * Return the registry of observed styles in insertion order.
     */
    public Map<String, IStyleData> getStyleRegistry() {
        return idToStyle;
    }

    /**
     * IStyleData → XSSFCellStyle，附带 quotePrefix=true，结果缓存。
     * FORCE_TEXT cells share a single quote-prefix variant per base style to
     * avoid blowing up the workbook's 64K style limit. The variant is built
     * from the same applyXxx helpers directly rather than cloning the base,
     * so a single empty-style force-text pool only contributes one extra
     * style object to the workbook.
     */
    public XSSFCellStyle toPoiStyleWithQuotePrefix(IStyleData s) {
        IStyleData src = s == null ? new IStyleData() : s;
        String key = styleIdOf(src) + "#qp";
        XSSFCellStyle cached = quotePrefixCache.get(key);
        if (cached != null) {
            return cached;
        }
        XSSFFont font = wb.createFont();
        applyFont(src, font);
        XSSFCellStyle style = wb.createCellStyle();
        applyBackground(src, style);
        applyBorders(src, style);
        applyAlignment(src, style);
        applyWrap(src, style);
        applyRotation(src, style);
        applyNumFmt(src, style);
        style.setFont(font);
        style.setQuotePrefixed(true);
        quotePrefixCache.put(key, style);
        return style;
    }

    // ============================================================
    // 写路径辅助方法 / Write-path helpers
    // ============================================================

    private void applyFont(IStyleData s, XSSFFont font) {
        if (s.getFf() != null) {
            font.setFontName(s.getFf());
        }
        if (s.getFs() != null) {
            font.setFontHeightInPoints(s.getFs().shortValue());
        }
        if (s.getBl() == BooleanNumber.TRUE) {
            font.setBold(true);
        }
        if (s.getIt() == BooleanNumber.TRUE) {
            font.setItalic(true);
        }
        if (s.getUl() != null && s.getUl().getS() == BooleanNumber.TRUE) {
            font.setUnderline(FontUnderline.SINGLE.getByteValue());
        }
        if (s.getSt() != null && s.getSt().getS() == BooleanNumber.TRUE) {
            font.setStrikeout(true);
        }
        if (s.getCl() != null && s.getCl().getRgb() != null) {
            font.setColor(new XSSFColor(ColorUtils.rgbHexToArgb(s.getCl().getRgb()), null));
        }
        if (s.getVa() != null) {
            int va = s.getVa();
            if (va == 2) {
                font.setTypeOffset(Font.SS_SUB);
            } else if (va == 3) {
                font.setTypeOffset(Font.SS_SUPER);
            } else {
                font.setTypeOffset(Font.SS_NONE);
            }
        }
    }

    private void applyBackground(IStyleData s, XSSFCellStyle style) {
        if (s.getBg() != null && s.getBg().getRgb() != null) {
            style.setFillForegroundColor(new XSSFColor(ColorUtils.rgbHexToArgb(s.getBg().getRgb()), null));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
    }

    private void applyBorders(IStyleData s, XSSFCellStyle style) {
        IBorderData bd = s.getBd();
        if (bd == null) {
            return;
        }
        applyBorderEdge(bd.getT(), style, Edge.TOP);
        applyBorderEdge(bd.getB(), style, Edge.BOTTOM);
        applyBorderEdge(bd.getL(), style, Edge.LEFT);
        applyBorderEdge(bd.getR(), style, Edge.RIGHT);
    }

    private void applyBorderEdge(IBorderStyleData edge, XSSFCellStyle style, Edge e) {
        if (edge == null || edge.getS() == null) {
            return;
        }
        int idx = edge.getS();
        BorderStyle[] values = BorderStyle.values();
        BorderStyle bs = (idx >= 0 && idx < values.length) ? values[idx] : BorderStyle.NONE;
        XSSFColor color = null;
        if (edge.getCl() != null && edge.getCl().getRgb() != null) {
            color = new XSSFColor(ColorUtils.rgbHexToArgb(edge.getCl().getRgb()), null);
        }
        switch (e) {
            case TOP:
                style.setBorderTop(bs);
                if (color != null) {
                    style.setTopBorderColor(color);
                }
                break;
            case BOTTOM:
                style.setBorderBottom(bs);
                if (color != null) {
                    style.setBottomBorderColor(color);
                }
                break;
            case LEFT:
                style.setBorderLeft(bs);
                if (color != null) {
                    style.setLeftBorderColor(color);
                }
                break;
            case RIGHT:
                style.setBorderRight(bs);
                if (color != null) {
                    style.setRightBorderColor(color);
                }
                break;
            default:
                break;
        }
    }

    private void applyAlignment(IStyleData s, XSSFCellStyle style) {
        if (s.getHt() != null) {
            int ht = s.getHt();
            if (ht == 1) {
                style.setAlignment(HorizontalAlignment.LEFT);
            } else if (ht == 2) {
                style.setAlignment(HorizontalAlignment.CENTER);
            } else if (ht == 3) {
                style.setAlignment(HorizontalAlignment.RIGHT);
            }
        }
        if (s.getVt() != null) {
            int vt = s.getVt();
            if (vt == 1) {
                style.setVerticalAlignment(VerticalAlignment.TOP);
            } else if (vt == 2) {
                style.setVerticalAlignment(VerticalAlignment.CENTER);
            } else if (vt == 3) {
                style.setVerticalAlignment(VerticalAlignment.BOTTOM);
            }
        }
    }

    private void applyWrap(IStyleData s, XSSFCellStyle style) {
        if (s.getTb() == null) {
            return;
        }
        int tb = s.getTb();
        if (tb == 2) {
            style.setShrinkToFit(true);
        } else if (tb == 3) {
            style.setWrapText(true);
        }
        // tb == 1 -> overflow, nothing to apply
    }

    private void applyRotation(IStyleData s, XSSFCellStyle style) {
        ITextRotation tr = s.getTr();
        if (tr == null) {
            return;
        }
        if (tr.getV() == BooleanNumber.TRUE) {
            style.setRotation((short) 255);
        } else if (tr.getA() != null) {
            style.setRotation(tr.getA().shortValue());
        }
    }

    private void applyNumFmt(IStyleData s, XSSFCellStyle style) {
        INumfmtLocal n = s.getN();
        if (n == null || n.getPattern() == null) {
            return;
        }
        style.setDataFormat(wb.createDataFormat().getFormat(n.getPattern()));
    }

    // ============================================================
    // 读路径辅助方法 / Read-path helpers
    // ============================================================

    private void readFont(XSSFCellStyle cs, IStyleData out) {
        XSSFFont font = cs.getFont();
        if (font == null) {
            return;
        }
        String name = font.getFontName();
        if (name != null) {
            out.setFf(name);
        }
        out.setFs((int) font.getFontHeightInPoints());
        if (font.getBold()) {
            out.setBl(BooleanNumber.TRUE);
        }
        if (font.getItalic()) {
            out.setIt(BooleanNumber.TRUE);
        }
        if (font.getUnderline() != (byte) 0) {
            out.setUl(new ITextDecoration().setS(BooleanNumber.TRUE));
        }
        if (font.getStrikeout()) {
            out.setSt(new ITextDecoration().setS(BooleanNumber.TRUE));
        }
        String fontRgb = xssfColorToHex(font.getXSSFColor());
        if (fontRgb != null) {
            out.setCl(new IColorStyle().setRgb(fontRgb));
        }
        short ts = font.getTypeOffset();
        if (ts == Font.SS_SUB) {
            out.setVa(2);
        } else if (ts == Font.SS_SUPER) {
            out.setVa(3);
        }
    }

    private void readBackground(XSSFCellStyle cs, IStyleData out) {
        if (cs.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
            return;
        }
        String rgb = xssfColorToHex(cs.getFillForegroundXSSFColor());
        if (rgb != null) {
            out.setBg(new IColorStyle().setRgb(rgb));
        }
    }

    private void readBorders(XSSFCellStyle cs, IStyleData out) {
        IBorderStyleData t = readEdge(cs.getBorderTop(), cs.getTopBorderXSSFColor());
        IBorderStyleData b = readEdge(cs.getBorderBottom(), cs.getBottomBorderXSSFColor());
        IBorderStyleData l = readEdge(cs.getBorderLeft(), cs.getLeftBorderXSSFColor());
        IBorderStyleData r = readEdge(cs.getBorderRight(), cs.getRightBorderXSSFColor());
        if (t == null && b == null && l == null && r == null) {
            return;
        }
        IBorderData bd = new IBorderData();
        if (t != null) bd.setT(t);
        if (b != null) bd.setB(b);
        if (l != null) bd.setL(l);
        if (r != null) bd.setR(r);
        out.setBd(bd);
    }

    private IBorderStyleData readEdge(BorderStyle bs, XSSFColor color) {
        if (bs == null || bs == BorderStyle.NONE) {
            return null;
        }
        IBorderStyleData e = new IBorderStyleData().setS(bs.ordinal());
        String rgb = xssfColorToHex(color);
        if (rgb != null) {
            e.setCl(new IColorStyle().setRgb(rgb));
        }
        return e;
    }

    private void readAlignment(XSSFCellStyle cs, IStyleData out) {
        HorizontalAlignment h = cs.getAlignment();
        if (h == HorizontalAlignment.LEFT) {
            out.setHt(1);
        } else if (h == HorizontalAlignment.CENTER) {
            out.setHt(2);
        } else if (h == HorizontalAlignment.RIGHT) {
            out.setHt(3);
        }
        VerticalAlignment v = cs.getVerticalAlignment();
        if (v == VerticalAlignment.TOP) {
            out.setVt(1);
        } else if (v == VerticalAlignment.CENTER) {
            out.setVt(2);
        } else if (v == VerticalAlignment.BOTTOM) {
            // BOTTOM 是 POI 默认值，若用户未显式设置则通常不希望出现在输出中；
            // 但为保证往返一致性，这里保留显式映射。
            // BOTTOM is POI's default; keep explicit mapping for round-trip fidelity.
            out.setVt(3);
        }
    }

    private void readWrap(XSSFCellStyle cs, IStyleData out) {
        if (cs.getWrapText()) {
            out.setTb(3);
        } else if (cs.getShrinkToFit()) {
            out.setTb(2);
        }
        // 否则视为默认(overflow)，不写出
    }

    private void readRotation(XSSFCellStyle cs, IStyleData out) {
        short r = cs.getRotation();
        if (r == 255) {
            out.setTr(new ITextRotation().setV(BooleanNumber.TRUE));
        } else if (r != 0) {
            out.setTr(new ITextRotation().setA((int) r));
        }
    }

    private void readNumFmt(XSSFCellStyle cs, IStyleData out) {
        String p = cs.getDataFormatString();
        if (p == null || p.isEmpty() || "General".equalsIgnoreCase(p)) {
            return;
        }
        out.setN(new INumfmtLocal().setPattern(p));
    }

    // ============================================================
    // 颜色 / 边框 helper
    // ============================================================

    private static String xssfColorToHex(XSSFColor color) {
        if (color == null) {
            return null;
        }
        byte[] argb = color.getARGB();
        if (argb == null || argb.length < 4) {
            return null;
        }
        return ColorUtils.argbToRgbHex(argb);
    }

    private enum Edge { TOP, BOTTOM, LEFT, RIGHT }
}
