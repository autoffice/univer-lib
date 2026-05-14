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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

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
    /** styleIdOf 计算结果缓存，键为 IStyleData 对象 / cache for styleIdOf computation. */
    private final Map<IStyleData, String> styleIdCache = new HashMap<>();

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
        // 先查缓存，避免重复计算 SHA-256 / check cache first to avoid redundant SHA-256 computation
        String cached = styleIdCache.get(src);
        if (cached != null) {
            return cached;
        }
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
            // 缓存计算结果 / cache the computed id
            styleIdCache.put(src, id);
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
            byte[] argb = ColorUtils.rgbHexToArgb(s.getCl().getRgb());
            if (argb != null) {
                font.setColor(new XSSFColor(argb, null));
            }
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
            byte[] argb = ColorUtils.rgbHexToArgb(s.getBg().getRgb());
            if (argb != null) {
                style.setFillForegroundColor(new XSSFColor(argb, null));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
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
        // Univer BorderStyleTypes 到 POI BorderStyle 的映射
        // Univer: NONE=0, THIN=1, HAIR=2, DOTTED=3, DASHED=4, DASH_DOT=5, DASH_DOT_DOT=6,
        //         DOUBLE=7, MEDIUM=8, MEDIUM_DASHED=9, MEDIUM_DASH_DOT=10, MEDIUM_DASH_DOT_DOT=11,
        //         SLANT_DASH_DOT=12, THICK=13
        // POI:    NONE=0, THIN=1, MEDIUM=2, DASHED=3, DOTTED=4, THICK=5, DOUBLE=6, HAIR=7,
        //         MEDIUM_DASHED=8, DASH_DOT=9, MEDIUM_DASH_DOT=10, DASH_DOT_DOT=11,
        //         MEDIUM_DASH_DOT_DOT=12, SLANTED_DASH_DOT=13
        BorderStyle bs = univerBorderStyleToPoi(edge.getS());
        XSSFColor color = null;
        if (edge.getCl() != null && edge.getCl().getRgb() != null) {
            byte[] argb = ColorUtils.rgbHexToArgb(edge.getCl().getRgb());
            if (argb != null) {
                color = new XSSFColor(argb, null);
            }
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
        // Univer 仅支持 SOLID_FOREGROUND 纯色填充；其他 PatternType（DOTS/BRICKS/SQUARES 等）
        // 在 IStyleData 中没有对应字段，丢弃。详见 docs/KNOWN_ISSUES.md。
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
            // POI 在 applyBorder 未显式设置时返回 NONE，但 OOXML 规范默认应用边框。
            // 回退到直接读取 CTBorder。
            // POI returns NONE when applyBorder is not explicitly set, but OOXML spec
            // defaults to applying the border. Fall back to reading CTBorder directly.
            readBordersFromCt(cs, out);
            return;
        }
        IBorderData bd = new IBorderData();
        if (t != null) bd.setT(t);
        if (b != null) bd.setB(b);
        if (l != null) bd.setL(l);
        if (r != null) bd.setR(r);
        out.setBd(bd);
    }

    private void readBordersFromCt(XSSFCellStyle cs, IStyleData out) {
        CTXf ctXf = cs.getCoreXf();
        if (ctXf == null || !ctXf.isSetBorderId()) {
            return;
        }
        long borderId = ctXf.getBorderId();
        if (borderId == 0) {
            return;
        }
        CTBorder ctBorder = wb.getStylesSource().getCTStylesheet()
            .getBorders().getBorderArray((int) borderId);
        if (ctBorder == null) {
            return;
        }
        IBorderStyleData t = readCtEdge(ctBorder.getTop());
        IBorderStyleData b = readCtEdge(ctBorder.getBottom());
        IBorderStyleData l = readCtEdge(ctBorder.getLeft());
        IBorderStyleData r = readCtEdge(ctBorder.getRight());
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

    private IBorderStyleData readCtEdge(CTBorderPr borderPr) {
        if (borderPr == null || !borderPr.isSetStyle()) {
            return null;
        }
        STBorderStyle.Enum style = borderPr.getStyle();
        if (style == null || style == STBorderStyle.NONE) {
            return null;
        }
        BorderStyle bs = ctBorderStyleToPoi(style);
        if (bs == null || bs == BorderStyle.NONE) {
            return null;
        }
        IBorderStyleData e = new IBorderStyleData().setS(poiBorderStyleToUniver(bs));
        CTColor ctColor = borderPr.getColor();
        if (ctColor != null) {
            String rgb = ctColorToHex(ctColor);
            if (rgb != null) {
                e.setCl(new IColorStyle().setRgb(rgb));
            }
        }
        return e;
    }

    private static BorderStyle ctBorderStyleToPoi(STBorderStyle.Enum st) {
        if (st == null) return BorderStyle.NONE;
        int v = st.intValue();
        switch (v) {
            case STBorderStyle.INT_NONE: return BorderStyle.NONE;
            case STBorderStyle.INT_THIN: return BorderStyle.THIN;
            case STBorderStyle.INT_MEDIUM: return BorderStyle.MEDIUM;
            case STBorderStyle.INT_DASHED: return BorderStyle.DASHED;
            case STBorderStyle.INT_DOTTED: return BorderStyle.DOTTED;
            case STBorderStyle.INT_THICK: return BorderStyle.THICK;
            case STBorderStyle.INT_DOUBLE: return BorderStyle.DOUBLE;
            case STBorderStyle.INT_HAIR: return BorderStyle.HAIR;
            case STBorderStyle.INT_MEDIUM_DASHED: return BorderStyle.MEDIUM_DASHED;
            case STBorderStyle.INT_DASH_DOT: return BorderStyle.DASH_DOT;
            case STBorderStyle.INT_MEDIUM_DASH_DOT: return BorderStyle.MEDIUM_DASH_DOT;
            case STBorderStyle.INT_DASH_DOT_DOT: return BorderStyle.DASH_DOT_DOT;
            case STBorderStyle.INT_MEDIUM_DASH_DOT_DOT: return BorderStyle.MEDIUM_DASH_DOT_DOT;
            case STBorderStyle.INT_SLANT_DASH_DOT: return BorderStyle.SLANTED_DASH_DOT;
            default: return BorderStyle.NONE;
        }
    }

    private static String ctColorToHex(CTColor ctColor) {
        if (ctColor == null) {
            return null;
        }
        byte[] rgb = ctColor.getRgb();
        if (rgb == null || rgb.length < 3) {
            return null;
        }
        if (rgb.length == 4) {
            return ColorUtils.argbToRgbHex(rgb);
        }
        return String.format("#%02x%02x%02x", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    private IBorderStyleData readEdge(BorderStyle bs, XSSFColor color) {
        if (bs == null || bs == BorderStyle.NONE) {
            return null;
        }
        IBorderStyleData e = new IBorderStyleData().setS(poiBorderStyleToUniver(bs));
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
            out.setVt(3);
        }

        // POI 在 applyAlignment 未显式设置时返回默认值（GENERAL / BOTTOM），
        // 但 OOXML 规范默认应用 alignment。从 CT 兜底读取。
        if (h == HorizontalAlignment.GENERAL || v == VerticalAlignment.BOTTOM) {
            readAlignmentFromCt(cs, out, h, v);
        }
    }

    private void readAlignmentFromCt(XSSFCellStyle cs, IStyleData out,
                                     HorizontalAlignment poiH, VerticalAlignment poiV) {
        CTXf ctXf = cs.getCoreXf();
        if (ctXf == null || !ctXf.isSetAlignment()) {
            return;
        }
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellAlignment ctAlign = ctXf.getAlignment();
        if (poiH == HorizontalAlignment.GENERAL && ctAlign.isSetHorizontal()) {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.Enum ha =
                ctAlign.getHorizontal();
            if (ha == org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.LEFT) {
                out.setHt(1);
            } else if (ha == org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.CENTER) {
                out.setHt(2);
            } else if (ha == org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment.RIGHT) {
                out.setHt(3);
            }
        }
        if (poiV == VerticalAlignment.BOTTOM && ctAlign.isSetVertical()) {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment.Enum va =
                ctAlign.getVertical();
            if (va == org.openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment.TOP) {
                out.setVt(1);
            } else if (va == org.openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment.CENTER) {
                out.setVt(2);
            }
            // BOTTOM 已经是默认值，不需要覆盖
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
    // 边框样式映射 / Border style mapping
    // ============================================================

    /**
     * Univer BorderStyleTypes 到 POI BorderStyle 的映射。
     * Map Univer BorderStyleTypes to POI BorderStyle.
     */
    private static BorderStyle univerBorderStyleToPoi(int univerStyle) {
        // Univer: NONE=0, THIN=1, HAIR=2, DOTTED=3, DASHED=4, DASH_DOT=5, DASH_DOT_DOT=6,
        //         DOUBLE=7, MEDIUM=8, MEDIUM_DASHED=9, MEDIUM_DASH_DOT=10, MEDIUM_DASH_DOT_DOT=11,
        //         SLANT_DASH_DOT=12, THICK=13
        // POI:    NONE=0, THIN=1, MEDIUM=2, DASHED=3, DOTTED=4, THICK=5, DOUBLE=6, HAIR=7,
        //         MEDIUM_DASHED=8, DASH_DOT=9, MEDIUM_DASH_DOT=10, DASH_DOT_DOT=11,
        //         MEDIUM_DASH_DOT_DOT=12, SLANTED_DASH_DOT=13
        switch (univerStyle) {
            case 0: return BorderStyle.NONE;
            case 1: return BorderStyle.THIN;
            case 2: return BorderStyle.HAIR;
            case 3: return BorderStyle.DOTTED;
            case 4: return BorderStyle.DASHED;
            case 5: return BorderStyle.DASH_DOT;
            case 6: return BorderStyle.DASH_DOT_DOT;
            case 7: return BorderStyle.DOUBLE;
            case 8: return BorderStyle.MEDIUM;
            case 9: return BorderStyle.MEDIUM_DASHED;
            case 10: return BorderStyle.MEDIUM_DASH_DOT;
            case 11: return BorderStyle.MEDIUM_DASH_DOT_DOT;
            case 12: return BorderStyle.SLANTED_DASH_DOT;
            case 13: return BorderStyle.THICK;
            default: return BorderStyle.NONE;
        }
    }

    /**
     * POI BorderStyle 到 Univer BorderStyleTypes 的映射。
     * Map POI BorderStyle to Univer BorderStyleTypes.
     */
    private static int poiBorderStyleToUniver(BorderStyle poiStyle) {
        if (poiStyle == null) return 0;
        // POI:    NONE=0, THIN=1, MEDIUM=2, DASHED=3, DOTTED=4, THICK=5, DOUBLE=6, HAIR=7,
        //         MEDIUM_DASHED=8, DASH_DOT=9, MEDIUM_DASH_DOT=10, DASH_DOT_DOT=11,
        //         MEDIUM_DASH_DOT_DOT=12, SLANTED_DASH_DOT=13
        // Univer: NONE=0, THIN=1, HAIR=2, DOTTED=3, DASHED=4, DASH_DOT=5, DASH_DOT_DOT=6,
        //         DOUBLE=7, MEDIUM=8, MEDIUM_DASHED=9, MEDIUM_DASH_DOT=10, MEDIUM_DASH_DOT_DOT=11,
        //         SLANT_DASH_DOT=12, THICK=13
        switch (poiStyle) {
            case NONE: return 0;
            case THIN: return 1;
            case MEDIUM: return 8;
            case DASHED: return 4;
            case DOTTED: return 3;
            case THICK: return 13;
            case DOUBLE: return 7;
            case HAIR: return 2;
            case MEDIUM_DASHED: return 9;
            case DASH_DOT: return 5;
            case MEDIUM_DASH_DOT: return 10;
            case DASH_DOT_DOT: return 6;
            case MEDIUM_DASH_DOT_DOT: return 11;
            case SLANTED_DASH_DOT: return 12;
            default: return 0;
        }
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
