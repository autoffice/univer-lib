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
package io.github.autoffice.univer.util;

/**
 * 颜色转换工具。
 * Color conversion utilities.
 */
public final class ColorUtils {
    private ColorUtils() {}

    /**
     * 将 #rrggbb 或 #aarrggbb 转为 ARGB 字节数组。
     * <p>输入为 null、空串、纯空白或格式非法时返回 {@code null}，调用方需判空后再使用。
     *
     * Convert hex color (#rrggbb or #aarrggbb) to ARGB byte array.
     * <p>Returns {@code null} for null, empty, blank or malformed input — callers must null-check.
     */
    public static byte[] rgbHexToArgb(String hex) {
        if (hex == null) {
            return null;
        }
        String s = hex.trim();
        if (s.isEmpty()) {
            return null;
        }
        if (s.startsWith("#")) {
            s = s.substring(1);
        }
        int a = 0xFF;
        int r;
        int g;
        int b;
        try {
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
                return null;
            }
        } catch (NumberFormatException e) {
            return null;
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
