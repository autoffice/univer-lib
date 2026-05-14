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
