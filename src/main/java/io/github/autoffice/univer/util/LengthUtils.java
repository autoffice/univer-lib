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
