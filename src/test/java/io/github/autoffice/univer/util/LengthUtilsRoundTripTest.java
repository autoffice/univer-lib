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

/**
 * LengthUtils 双向换算覆盖。
 * Round-trips px↔pt 与 px↔chars 以保证两向函数都跑到。
 */
class LengthUtilsRoundTripTest {

    @Test
    void should_convert_pt_back_to_px() {
        double px = 96.0;
        double pt = LengthUtils.pxToPoints(px);
        assertThat(pt).isEqualTo(72.0);
        assertThat(LengthUtils.pointsToPx(pt)).isEqualTo(px);
    }

    @Test
    void should_convert_chars_back_to_px() {
        double px = 70.0;
        double chars = LengthUtils.pxToChars(px);
        assertThat(chars).isEqualTo(10.0);
        assertThat(LengthUtils.charsToPx(chars)).isEqualTo(px);
    }
}
