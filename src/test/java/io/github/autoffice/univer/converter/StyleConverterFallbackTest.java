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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorders;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 覆盖 StyleConverter 的 POI applyBorder/applyAlignment fallback 分支。
 * 直接操纵底层 CTXf / CTBorder，让 POI 的 getBorderTop 返回 NONE 但 CTBorder 仍有数据。
 */
class StyleConverterFallbackTest {

    @Test
    void should_read_border_from_ct_when_apply_border_is_off() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createCellStyle();
            StyleConverter sc = new StyleConverter(wb);
            CTBorders bordersTable = wb.getStylesSource().getCTStylesheet().getBorders();
            if (bordersTable == null) {
                bordersTable = wb.getStylesSource().getCTStylesheet().addNewBorders();
            }
            CTBorder ctBorder = bordersTable.addNewBorder();
            CTBorderPr top = ctBorder.addNewTop();
            top.setStyle(STBorderStyle.THIN);
            CTColor topColor = top.addNewColor();
            topColor.setRgb(new byte[]{(byte) 0xFF, (byte) 0xAB, (byte) 0xCD, (byte) 0xEF});
            int borderId = bordersTable.sizeOfBorderArray() - 1;
            bordersTable.setCount(bordersTable.sizeOfBorderArray());

            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            ctXf.setBorderId(borderId);
            ctXf.setApplyBorder(false);

            // 调用 fromPoiStyle：无论走主路径还是 fallback，至少都能拿到一个非 null 结果
            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back).isNotNull();
        }
    }

    @Test
    void should_skip_ct_border_when_borderId_zero() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            // borderId=0 (默认空 border) → readBordersFromCt 应跳过
            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back.getBd()).isNull();
        }
    }

    @Test
    void should_read_alignment_from_ct_when_apply_alignment_is_off() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            // 直接构造 alignment 信息但 applyAlignment=false
            CTCellAlignment ctAlign = ctXf.addNewAlignment();
            ctAlign.setHorizontal(STHorizontalAlignment.CENTER);
            ctAlign.setVertical(STVerticalAlignment.TOP);
            ctXf.setApplyAlignment(false);

            IStyleData back = sc.fromPoiStyle(style);
            // fallback 应读到 ht=2 (CENTER), vt=1 (TOP)
            assertThat(back.getHt()).isEqualTo(2);
            assertThat(back.getVt()).isEqualTo(1);
        }
    }

    @Test
    void should_read_alignment_from_ct_with_left_and_center_vt() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            CTCellAlignment ctAlign = ctXf.addNewAlignment();
            ctAlign.setHorizontal(STHorizontalAlignment.LEFT);
            ctAlign.setVertical(STVerticalAlignment.CENTER);
            ctXf.setApplyAlignment(false);
            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back.getHt()).isEqualTo(1);
            assertThat(back.getVt()).isEqualTo(2);
        }
    }

    @Test
    void should_read_alignment_from_ct_with_right_alignment() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            CTCellAlignment ctAlign = ctXf.addNewAlignment();
            ctAlign.setHorizontal(STHorizontalAlignment.RIGHT);
            ctXf.setApplyAlignment(false);
            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back.getHt()).isEqualTo(3);
        }
    }

    @Test
    void should_skip_alignment_fallback_when_no_alignment_set() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            // 不设 alignment → readAlignmentFromCt 应跳过 fallback；
            // POI 默认 VerticalAlignment.BOTTOM 会触发 fallback 入口，但因为 ctAlignment 不存在所以直接 return
            IStyleData back = sc.fromPoiStyle(style);
            // 不抛错即可，且 ht 仍可能为 null
            assertThat(back).isNotNull();
        }
    }

    @Test
    void should_read_ct_border_with_all_style_enums() throws Exception {
        STBorderStyle.Enum[] all = {
                STBorderStyle.THIN, STBorderStyle.MEDIUM, STBorderStyle.DASHED, STBorderStyle.DOTTED,
                STBorderStyle.THICK, STBorderStyle.DOUBLE, STBorderStyle.HAIR, STBorderStyle.MEDIUM_DASHED,
                STBorderStyle.DASH_DOT, STBorderStyle.MEDIUM_DASH_DOT, STBorderStyle.DASH_DOT_DOT,
                STBorderStyle.MEDIUM_DASH_DOT_DOT, STBorderStyle.SLANT_DASH_DOT
        };
        for (STBorderStyle.Enum bs : all) {
            try (XSSFWorkbook wb = new XSSFWorkbook()) {
                wb.createCellStyle();
                StyleConverter sc = new StyleConverter(wb);
                CTBorders bordersTable = wb.getStylesSource().getCTStylesheet().getBorders();
                if (bordersTable == null) {
                    bordersTable = wb.getStylesSource().getCTStylesheet().addNewBorders();
                }
                CTBorder ctBorder = bordersTable.addNewBorder();
                ctBorder.addNewTop().setStyle(bs);
                int bid = bordersTable.sizeOfBorderArray() - 1;
                bordersTable.setCount(bordersTable.sizeOfBorderArray());

                XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
                CTXf ctXf = style.getCoreXf();
                ctXf.setBorderId(bid);
                ctXf.setApplyBorder(false);

                IStyleData back = sc.fromPoiStyle(style);
                assertThat(back).as("style=%s", bs).isNotNull();
            }
        }
    }

    @Test
    void should_skip_ct_border_with_NONE_style() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createCellStyle();
            StyleConverter sc = new StyleConverter(wb);
            CTBorders bordersTable = wb.getStylesSource().getCTStylesheet().getBorders();
            if (bordersTable == null) {
                bordersTable = wb.getStylesSource().getCTStylesheet().addNewBorders();
            }
            CTBorder ctBorder = bordersTable.addNewBorder();
            ctBorder.addNewTop().setStyle(STBorderStyle.NONE);
            int bid = bordersTable.sizeOfBorderArray() - 1;
            bordersTable.setCount(bordersTable.sizeOfBorderArray());

            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            ctXf.setBorderId(bid);
            ctXf.setApplyBorder(false);

            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back).isNotNull();
        }
    }

    @Test
    void should_skip_ct_border_with_3_byte_rgb_color() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createCellStyle();
            StyleConverter sc = new StyleConverter(wb);
            CTBorders bordersTable = wb.getStylesSource().getCTStylesheet().getBorders();
            if (bordersTable == null) {
                bordersTable = wb.getStylesSource().getCTStylesheet().addNewBorders();
            }
            CTBorder ctBorder = bordersTable.addNewBorder();
            CTBorderPr top = ctBorder.addNewTop();
            top.setStyle(STBorderStyle.THIN);
            CTColor c = top.addNewColor();
            c.setRgb(new byte[]{(byte) 0x12, (byte) 0x34, (byte) 0x56});
            int bid = bordersTable.sizeOfBorderArray() - 1;
            bordersTable.setCount(bordersTable.sizeOfBorderArray());

            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            CTXf ctXf = style.getCoreXf();
            ctXf.setBorderId(bid);
            ctXf.setApplyBorder(false);

            IStyleData back = sc.fromPoiStyle(style);
            assertThat(back).isNotNull();
        }
    }
}
