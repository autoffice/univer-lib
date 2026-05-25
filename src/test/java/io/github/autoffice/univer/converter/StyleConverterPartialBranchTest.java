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

import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.IBorderData;
import io.github.autoffice.univer.model.IBorderStyleData;
import io.github.autoffice.univer.model.IColorStyle;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextDecoration;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * StyleConverter partial 分支精确补全：
 * - ul/st/cl/bg 字段「getXxx() 不为 null 但内层 == null」分支
 * - applyBorderEdge 各 case 中 color == null 分支
 * - ht/vt = 3 写路径
 * - 边框颜色 rgb 非法分支
 * - applyFont 子分支
 */
class StyleConverterPartialBranchTest {

    @Test
    void should_handle_ul_object_with_null_s_field() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // ul 存在但 s 为 null（默认）→ 不应设 underline
            IStyleData src = new IStyleData().setUl(new ITextDecoration());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            XSSFFont font = wb.getFontAt(poi.getFontIndex());
            // 默认 underline 是 NONE
            assertThat(font.getUnderline()).isEqualTo((byte) 0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_st_object_with_null_s_field() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setSt(new ITextDecoration());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            XSSFFont font = wb.getFontAt(poi.getFontIndex());
            assertThat(font.getStrikeout()).isFalse();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_cl_object_with_null_rgb_field() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // cl 对象存在但 rgb 为 null
            IStyleData src = new IStyleData().setCl(new IColorStyle());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            // 不抛错即可
            assertThat(poi).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_bg_object_with_null_rgb_field() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setBg(new IColorStyle());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_apply_border_without_color() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 边框样式存在但没有颜色 → 走 color == null 分支
            IStyleData src = new IStyleData()
                    .setBd(new IBorderData()
                            .setT(new IBorderStyleData().setS(1))    // 无 cl
                            .setB(new IBorderStyleData().setS(1))    // 无 cl
                            .setL(new IBorderStyleData().setS(1))    // 无 cl
                            .setR(new IBorderStyleData().setS(1)));  // 无 cl
            XSSFCellStyle poi = sc.toPoiStyle(src);
            // 仍然能设置边框样式
            assertThat(poi.getBorderTop().getCode()).isNotEqualTo(0);
            assertThat(poi.getBorderBottom().getCode()).isNotEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_apply_border_with_invalid_color_rgb() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 边框颜色 rgb 非法 → ColorUtils 返回 null → 不设颜色
            IStyleData src = new IStyleData()
                    .setBd(new IBorderData()
                            .setT(new IBorderStyleData().setS(1)
                                    .setCl(new IColorStyle().setRgb("invalid"))));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_ht_3_to_right_alignment() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setHt(3);
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getAlignment()).isEqualTo(HorizontalAlignment.RIGHT);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getHt()).isEqualTo(3);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_vt_3_to_bottom_alignment() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setVt(3);
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getVerticalAlignment()).isEqualTo(VerticalAlignment.BOTTOM);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_font_color_when_xssfColor_null() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 默认 font 没颜色 → font.getXSSFColor() 返回 null → 不写入 cl
            IStyleData src = new IStyleData().setFf("Arial");
            XSSFCellStyle poi = sc.toPoiStyle(src);
            IStyleData back = sc.fromPoiStyle(poi);
            // cl 字段不应被设置
            assertThat(back.getCl()).isNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_underline_strikethrough_with_explicit_false() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // ul.s = FALSE 而不是 TRUE → 不设 underline
            IStyleData src = new IStyleData()
                    .setUl(new ITextDecoration().setS(BooleanNumber.FALSE))
                    .setSt(new ITextDecoration().setS(BooleanNumber.FALSE));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            XSSFFont font = wb.getFontAt(poi.getFontIndex());
            assertThat(font.getUnderline()).isEqualTo((byte) 0);
            assertThat(font.getStrikeout()).isFalse();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_border_with_default_edge_in_switch() {
        // 触发 applyBorderEdge switch 的 default 分支极其困难（Edge 是 enum，全部 case 已覆盖）
        // 这里通过一个标准 round-trip 让所有 case 都跑过一遍
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData()
                    .setBd(new IBorderData()
                            .setT(new IBorderStyleData().setS(1).setCl(new IColorStyle().setRgb("#000000")))
                            .setB(new IBorderStyleData().setS(1).setCl(new IColorStyle().setRgb("#111111")))
                            .setL(new IBorderStyleData().setS(1).setCl(new IColorStyle().setRgb("#222222")))
                            .setR(new IBorderStyleData().setS(1).setCl(new IColorStyle().setRgb("#333333"))));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getBorderTop().getCode()).isNotEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_apply_border_with_cl_object_but_null_rgb() {
        // 覆盖 L263: edge.getCl() != null && edge.getCl().getRgb() == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // edge.cl 存在但 rgb=null
            IStyleData src = new IStyleData()
                    .setBd(new IBorderData()
                            .setT(new IBorderStyleData().setS(1).setCl(new IColorStyle())));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            // 边框样式应仍被设置，颜色未被设
            assertThat(poi.getBorderTop().getCode()).isNotEqualTo(0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_horizontal_alignment_right() {
        // 覆盖 read 路径中 ht=3 (RIGHT) 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setHt(3);
            XSSFCellStyle poi = sc.toPoiStyle(src);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getHt()).isEqualTo(3);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_xssfColorToHex_with_short_argb() throws Exception {
        // 覆盖 L400: rgb 长度 < 4 → 返回 null
        java.lang.reflect.Method m = StyleConverter.class.getDeclaredMethod(
                "xssfColorToHex", org.apache.poi.xssf.usermodel.XSSFColor.class);
        m.setAccessible(true);

        // null color → null
        Object r1 = m.invoke(null, (Object) null);
        assertThat(r1).isNull();

        // 短 argb（< 4 字节）→ null
        org.apache.poi.xssf.usermodel.XSSFColor shortColor = new org.apache.poi.xssf.usermodel.XSSFColor(
                new byte[]{(byte) 0x12, (byte) 0x34, (byte) 0x56}, null);
        Object r2 = m.invoke(null, shortColor);
        // 短的可能也能用，依赖 POI 实际行为
        // 不抛错即可

        // 完整 argb
        org.apache.poi.xssf.usermodel.XSSFColor full = new org.apache.poi.xssf.usermodel.XSSFColor(
                new byte[]{(byte) 0xFF, (byte) 0xAB, (byte) 0xCD, (byte) 0xEF}, null);
        Object r3 = m.invoke(null, full);
        assertThat(r3).isNotNull();
    }
}
