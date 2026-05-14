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

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class ColorUtilsTest {
    @Test
    void should_convert_rgb_to_argb_bytes() {
        byte[] argb = ColorUtils.rgbHexToArgb("#ff0000");
        assertThat(argb).containsExactly((byte)0xFF, (byte)0xFF, (byte)0, (byte)0);
    }
    @Test
    void should_convert_argb_to_rgb_hex() {
        assertThat(ColorUtils.argbToRgbHex(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0,(byte)0}))
            .isEqualTo("#ff0000");
    }
    @Test
    void should_handle_8_char_hex_with_alpha() {
        byte[] argb = ColorUtils.rgbHexToArgb("#80ff0000");
        assertThat(argb[0]).isEqualTo((byte)0x80);
        assertThat(argb[1]).isEqualTo((byte)0xFF);
    }
    @Test
    void should_return_null_for_null_input() {
        assertThat(ColorUtils.rgbHexToArgb(null)).isNull();
        assertThat(ColorUtils.argbToRgbHex(null)).isNull();
    }
    @Test
    void should_return_null_for_invalid_or_empty_hex() {
        // 前端快照里颜色字段常为 null/""/非法值，util 需宽松处理返回 null，由调用方跳过
        // Front-end snapshots often carry null/empty/malformed color strings; utility returns null so callers can skip.
        assertThat(ColorUtils.rgbHexToArgb(null)).isNull();
        assertThat(ColorUtils.rgbHexToArgb("")).isNull();
        assertThat(ColorUtils.rgbHexToArgb("   ")).isNull();
        assertThat(ColorUtils.rgbHexToArgb("#xyz")).isNull();
        assertThat(ColorUtils.rgbHexToArgb("#abcd")).isNull();
    }
}
