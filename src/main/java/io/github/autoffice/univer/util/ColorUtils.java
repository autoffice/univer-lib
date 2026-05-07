package io.github.autoffice.univer.util;

/**
 * 颜色转换工具。
 * Color conversion utilities.
 */
public final class ColorUtils {
    private ColorUtils() {}

    /**
     * 将 #rrggbb 或 #aarrggbb 转为 ARGB 字节数组。
     * Convert hex color (#rrggbb or #aarrggbb) to ARGB byte array.
     */
    public static byte[] rgbHexToArgb(String hex) {
        if (hex == null) return null;
        String s = hex.startsWith("#") ? hex.substring(1) : hex;
        int a = 0xFF, r, g, b;
        if (s.length() == 6) {
            r = Integer.parseInt(s.substring(0, 2), 16);
            g = Integer.parseInt(s.substring(2, 4), 16);
            b = Integer.parseInt(s.substring(4, 6), 16);
        } else if (s.length() == 8) {
            a = Integer.parseInt(s.substring(0, 2), 16);
            r = Integer.parseInt(s.substring(2, 4), 16);
            g = Integer.parseInt(s.substring(4, 6), 16);
            b = Integer.parseInt(s.substring(6, 8), 16);
        } else {
            throw new IllegalArgumentException("Invalid hex color: " + hex);
        }
        return new byte[]{(byte) a, (byte) r, (byte) g, (byte) b};
    }

    /**
     * 将 ARGB 字节数组转为 #rrggbb（丢弃透明度）。
     * Convert ARGB bytes to #rrggbb hex (alpha discarded).
     */
    public static String argbToRgbHex(byte[] argb) {
        if (argb == null || argb.length < 4) return null;
        return String.format("#%02x%02x%02x", argb[1] & 0xFF, argb[2] & 0xFF, argb[3] & 0xFF);
    }
}
