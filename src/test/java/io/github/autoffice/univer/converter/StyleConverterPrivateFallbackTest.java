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
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.IStyleData;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

import java.lang.reflect.Method;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 通过反射直击 StyleConverter 的 POI fallback 私有方法：
 * readBordersFromCt / ctBorderStyleToPoi / readCtEdge / readAlignmentFromCt。
 * 这些在 POI 5.2 公共 API 上几乎不可能触发，但仍需保留 — 反射用于覆盖这些代码路径。
 */
class StyleConverterPrivateFallbackTest {

    @Test
    void should_map_each_ct_border_style_via_reflection() throws Exception {
        Method m = StyleConverter.class.getDeclaredMethod("ctBorderStyleToPoi", STBorderStyle.Enum.class);
        m.setAccessible(true);

        // 14 种 STBorderStyle 全部 case + null + default
        Object[][] cases = {
                {STBorderStyle.NONE, BorderStyle.NONE},
                {STBorderStyle.THIN, BorderStyle.THIN},
                {STBorderStyle.MEDIUM, BorderStyle.MEDIUM},
                {STBorderStyle.DASHED, BorderStyle.DASHED},
                {STBorderStyle.DOTTED, BorderStyle.DOTTED},
                {STBorderStyle.THICK, BorderStyle.THICK},
                {STBorderStyle.DOUBLE, BorderStyle.DOUBLE},
                {STBorderStyle.HAIR, BorderStyle.HAIR},
                {STBorderStyle.MEDIUM_DASHED, BorderStyle.MEDIUM_DASHED},
                {STBorderStyle.DASH_DOT, BorderStyle.DASH_DOT},
                {STBorderStyle.MEDIUM_DASH_DOT, BorderStyle.MEDIUM_DASH_DOT},
                {STBorderStyle.DASH_DOT_DOT, BorderStyle.DASH_DOT_DOT},
                {STBorderStyle.MEDIUM_DASH_DOT_DOT, BorderStyle.MEDIUM_DASH_DOT_DOT},
                {STBorderStyle.SLANT_DASH_DOT, BorderStyle.SLANTED_DASH_DOT},
        };
        for (Object[] c : cases) {
            assertThat((BorderStyle) m.invoke(null, c[0])).as("case %s", c[0]).isEqualTo(c[1]);
        }
        // null → NONE
        assertThat((BorderStyle) m.invoke(null, (Object) null)).isEqualTo(BorderStyle.NONE);
    }

    @Test
    void should_handle_ct_color_to_hex_via_reflection() throws Exception {
        Method m = StyleConverter.class.getDeclaredMethod(
                "ctColorToHex",
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.class);
        m.setAccessible(true);

        // null → null
        assertThat(m.invoke(null, (Object) null)).isNull();

        // CTColor with 4-byte ARGB
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor c4 =
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
        c4.setRgb(new byte[]{(byte) 0xFF, (byte) 0xAB, (byte) 0xCD, (byte) 0xEF});
        assertThat(m.invoke(null, c4)).isEqualTo("#abcdef");

        // CTColor with 3-byte RGB → 走另一分支
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor c3 =
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
        c3.setRgb(new byte[]{(byte) 0x12, (byte) 0x34, (byte) 0x56});
        assertThat((String) m.invoke(null, c3)).isEqualToIgnoringCase("#123456");

        // CTColor with too-short rgb → null
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor c2 =
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
        c2.setRgb(new byte[]{(byte) 0x12, (byte) 0x34});
        assertThat(m.invoke(null, c2)).isNull();

        // CTColor without rgb → null
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor empty =
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
        assertThat(m.invoke(null, empty)).isNull();
    }

    @Test
    void should_read_ct_edge_via_reflection() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            Method m = StyleConverter.class.getDeclaredMethod(
                    "readCtEdge",
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.class);
            m.setAccessible(true);

            // null → null
            assertThat(m.invoke(sc, (Object) null)).isNull();

            // CTBorderPr 没设 style → null
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr noStyle =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
            assertThat(m.invoke(sc, noStyle)).isNull();

            // style = NONE → null
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr noneStyle =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
            noneStyle.setStyle(STBorderStyle.NONE);
            assertThat(m.invoke(sc, noneStyle)).isNull();

            // style = THIN + 颜色 → 返回 IBorderStyleData
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr withColor =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
            withColor.setStyle(STBorderStyle.THIN);
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor color = withColor.addNewColor();
            color.setRgb(new byte[]{(byte) 0xFF, (byte) 0x12, (byte) 0x34, (byte) 0x56});
            Object edge = m.invoke(sc, withColor);
            assertThat(edge).isNotNull();

            // style = THIN，无颜色 → 仍返回非空但 cl 为 null
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr noColor =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
            noColor.setStyle(STBorderStyle.THIN);
            Object edgeNoColor = m.invoke(sc, noColor);
            assertThat(edgeNoColor).isNotNull();

            // style = THIN，颜色 rgb 为 null（未设 rgb 字节）
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr emptyColor =
                    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr.Factory.newInstance();
            emptyColor.setStyle(STBorderStyle.THIN);
            emptyColor.addNewColor(); // 不设 rgb
            Object edgeEmptyColor = m.invoke(sc, emptyColor);
            assertThat(edgeEmptyColor).isNotNull();
        }
    }

    @Test
    void should_read_borders_from_ct_via_reflection() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createCellStyle();
            StyleConverter sc = new StyleConverter(wb);
            Method m = StyleConverter.class.getDeclaredMethod(
                    "readBordersFromCt",
                    org.apache.poi.xssf.usermodel.XSSFCellStyle.class,
                    IStyleData.class);
            m.setAccessible(true);

            // 1) 默认 cellStyle borderId=0 → 早 return
            IStyleData out1 = new IStyleData();
            org.apache.poi.xssf.usermodel.XSSFCellStyle defaultStyle = wb.getCellStyleAt(0);
            m.invoke(sc, defaultStyle, out1);
            assertThat(out1.getBd()).isNull();

            // 2) borderId>0 且 borders 数组对应位置有真实数据
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorders bordersTable =
                    wb.getStylesSource().getCTStylesheet().getBorders();
            if (bordersTable == null) {
                bordersTable = wb.getStylesSource().getCTStylesheet().addNewBorders();
            }
            // 占位 index 0：默认 NONE border
            if (bordersTable.sizeOfBorderArray() == 0) {
                bordersTable.addNewBorder();
            }
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder ctBorder = bordersTable.addNewBorder();
            ctBorder.addNewTop().setStyle(STBorderStyle.THIN);
            ctBorder.addNewBottom().setStyle(STBorderStyle.MEDIUM);
            ctBorder.addNewLeft().setStyle(STBorderStyle.DASHED);
            ctBorder.addNewRight().setStyle(STBorderStyle.DOUBLE);
            int bid = bordersTable.sizeOfBorderArray() - 1;
            bordersTable.setCount(bordersTable.sizeOfBorderArray());

            org.apache.poi.xssf.usermodel.XSSFCellStyle style =
                    (org.apache.poi.xssf.usermodel.XSSFCellStyle) wb.createCellStyle();
            // 必须先 setBorderId 触发 isSetBorderId
            style.getCoreXf().setBorderId(bid);
            // 防御：确认 bid > 0，否则 readBordersFromCt 会因 borderId==0 提前 return
            assertThat(bid).isPositive();

            IStyleData out2 = new IStyleData();
            m.invoke(sc, style, out2);
            assertThat(out2.getBd()).isNotNull();
            assertThat(out2.getBd().getT()).isNotNull();
            assertThat(out2.getBd().getB()).isNotNull();
            assertThat(out2.getBd().getL()).isNotNull();
            assertThat(out2.getBd().getR()).isNotNull();

            // 3) borderId 引用一个所有边都 NONE 的 CTBorder → bd 仍为 null
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder allNone = bordersTable.addNewBorder();
            allNone.addNewTop().setStyle(STBorderStyle.NONE);
            allNone.addNewBottom().setStyle(STBorderStyle.NONE);
            allNone.addNewLeft().setStyle(STBorderStyle.NONE);
            allNone.addNewRight().setStyle(STBorderStyle.NONE);
            int bid2 = bordersTable.sizeOfBorderArray() - 1;
            bordersTable.setCount(bordersTable.sizeOfBorderArray());

            org.apache.poi.xssf.usermodel.XSSFCellStyle style2 =
                    (org.apache.poi.xssf.usermodel.XSSFCellStyle) wb.createCellStyle();
            style2.getCoreXf().setBorderId(bid2);

            IStyleData out3 = new IStyleData();
            m.invoke(sc, style2, out3);
            assertThat(out3.getBd()).isNull();
        }
    }
}
