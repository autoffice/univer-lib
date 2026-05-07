package io.github.autoffice.univer.util;

import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;

class LengthUtilsTest {
    @Test
    void should_convert_px_to_points() {
        assertThat(LengthUtils.pxToPoints(96.0)).isCloseTo(72.0, within(0.01));
    }
    @Test
    void should_convert_points_to_px() {
        assertThat(LengthUtils.pointsToPx(72.0)).isCloseTo(96.0, within(0.01));
    }
    @Test
    void should_convert_px_to_chars() {
        assertThat(LengthUtils.pxToChars(70.0)).isCloseTo(10.0, within(0.01));
    }
    @Test
    void should_convert_chars_to_px() {
        assertThat(LengthUtils.charsToPx(10.0)).isCloseTo(70.0, within(0.01));
    }
}
