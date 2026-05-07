package io.github.autoffice.univer.util;

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

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
    void should_throw_on_invalid_hex() {
        assertThatThrownBy(() -> ColorUtils.rgbHexToArgb("#xyz"))
            .isInstanceOf(IllegalArgumentException.class);
    }
}
