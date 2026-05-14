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
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class LeafPojoJsonTest {
    private final ObjectMapper m = JsonMapper.get();

    @Test
    void should_roundtrip_color_style() throws Exception {
        IColorStyle c = new IColorStyle().setRgb("#ff0000");
        String json = m.writeValueAsString(c);
        assertThat(json).isEqualTo("{\"rgb\":\"#ff0000\"}");
        assertThat(m.readValue(json, IColorStyle.class)).isEqualTo(c);
    }

    @Test
    void should_roundtrip_border_data() throws Exception {
        IBorderData bd = new IBorderData().setT(new IBorderStyleData().setS(1)
            .setCl(new IColorStyle().setRgb("#000000")));
        String json = m.writeValueAsString(bd);
        assertThat(m.readValue(json, IBorderData.class)).isEqualTo(bd);
    }

    @Test
    void should_roundtrip_range() throws Exception {
        IRange r = new IRange().setStartRow(0).setStartColumn(0).setEndRow(2).setEndColumn(3);
        String json = m.writeValueAsString(r);
        assertThat(m.readValue(json, IRange.class)).isEqualTo(r);
    }

    @Test
    void should_roundtrip_row_and_column() throws Exception {
        IRowData row = new IRowData().setH(20.0).setHd(BooleanNumber.TRUE);
        IColumnData col = new IColumnData().setW(100.0);
        assertThat(m.readValue(m.writeValueAsString(row), IRowData.class)).isEqualTo(row);
        assertThat(m.readValue(m.writeValueAsString(col), IColumnData.class)).isEqualTo(col);
    }

    @Test
    void should_roundtrip_freeze() throws Exception {
        IFreeze f = new IFreeze().setStartRow(1).setStartColumn(1).setXSplit(1).setYSplit(1);
        assertThat(m.readValue(m.writeValueAsString(f), IFreeze.class)).isEqualTo(f);
    }
}
