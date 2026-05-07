package io.github.autoffice.univer.util;

/**
 * 长度换算工具（近似）。
 * Length conversion utilities (approximate).
 */
public final class LengthUtils {
    private static final double PX_PER_INCH = 96.0;
    private static final double PT_PER_INCH = 72.0;
    private static final double PX_PER_CHAR = 7.0;

    private LengthUtils() {}

    /** px → pt。/ Convert px to points. */
    public static double pxToPoints(double px) { return px * PT_PER_INCH / PX_PER_INCH; }
    /** pt → px。/ Convert points to px. */
    public static double pointsToPx(double pt) { return pt * PX_PER_INCH / PT_PER_INCH; }
    /** px → 字符数（Excel 列宽单位）。/ Convert px to chars (Excel column-width unit). */
    public static double pxToChars(double px) { return px / PX_PER_CHAR; }
    /** 字符数 → px。/ Convert chars to px. */
    public static double charsToPx(double chars) { return chars * PX_PER_CHAR; }
}
