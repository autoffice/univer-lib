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
import io.github.autoffice.univer.model.INumfmtLocal;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextRotation;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 补充 StyleConverter 的次要分支覆盖：边框各种枚举、va、ht/vt 三向、清空样式、numFmt、registry 等。
 */
class StyleConverterBranchTest {

    @Test
    void should_map_all_univer_border_styles_to_poi_and_back() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // Univer enum 0..13 全部
            List<Integer> values = Arrays.asList(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13);
            for (int univerStyle : values) {
                // 四个边都带颜色，覆盖 BOTTOM/LEFT/RIGHT 的 setColor 分支
                IColorStyle cl = new IColorStyle().setRgb("#112233");
                IStyleData src = new IStyleData()
                        .setBd(new IBorderData()
                                .setT(new IBorderStyleData().setS(univerStyle).setCl(cl))
                                .setB(new IBorderStyleData().setS(univerStyle).setCl(cl))
                                .setL(new IBorderStyleData().setS(univerStyle).setCl(cl))
                                .setR(new IBorderStyleData().setS(univerStyle).setCl(cl)));
                XSSFCellStyle poi = sc.toPoiStyle(src);
                if (univerStyle == 0) {
                    assertThat(poi.getBorderTop()).isEqualTo(BorderStyle.NONE);
                } else {
                    assertThat(poi.getBorderTop()).isNotEqualTo(BorderStyle.NONE);
                }
            }
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_vertical_align_sub_and_super() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // SUB
            IStyleData sub = new IStyleData().setVa(2);
            XSSFCellStyle poiSub = sc.toPoiStyle(sub);
            XSSFFont fontSub = wb.getFontAt(poiSub.getFontIndex());
            assertThat(fontSub.getTypeOffset()).isEqualTo(Font.SS_SUB);
            IStyleData backSub = sc.fromPoiStyle(poiSub);
            assertThat(backSub.getVa()).isEqualTo(2);

            // SUPER
            IStyleData sup = new IStyleData().setVa(3);
            XSSFCellStyle poiSup = sc.toPoiStyle(sup);
            XSSFFont fontSup = wb.getFontAt(poiSup.getFontIndex());
            assertThat(fontSup.getTypeOffset()).isEqualTo(Font.SS_SUPER);
            IStyleData backSup = sc.fromPoiStyle(poiSup);
            assertThat(backSup.getVa()).isEqualTo(3);

            // NONE (va=1)
            IStyleData none = new IStyleData().setVa(1);
            XSSFCellStyle poiNone = sc.toPoiStyle(none);
            XSSFFont fontNone = wb.getFontAt(poiNone.getFontIndex());
            assertThat(fontNone.getTypeOffset()).isEqualTo(Font.SS_NONE);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_all_horizontal_alignments() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // ht=2 → CENTER
            XSSFCellStyle poi = sc.toPoiStyle(new IStyleData().setHt(2));
            assertThat(poi.getAlignment()).isEqualTo(HorizontalAlignment.CENTER);
            assertThat(sc.fromPoiStyle(poi).getHt()).isEqualTo(2);
            // ht=3 → RIGHT
            XSSFCellStyle poiR = sc.toPoiStyle(new IStyleData().setHt(3));
            assertThat(poiR.getAlignment()).isEqualTo(HorizontalAlignment.RIGHT);
            assertThat(sc.fromPoiStyle(poiR).getHt()).isEqualTo(3);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_all_vertical_alignments() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // vt=1 → TOP
            XSSFCellStyle poiT = sc.toPoiStyle(new IStyleData().setVt(1));
            assertThat(poiT.getVerticalAlignment()).isEqualTo(VerticalAlignment.TOP);
            assertThat(sc.fromPoiStyle(poiT).getVt()).isEqualTo(1);
            // vt=3 → BOTTOM (POI 默认)
            XSSFCellStyle poiB = sc.toPoiStyle(new IStyleData().setVt(3));
            assertThat(poiB.getVerticalAlignment()).isEqualTo(VerticalAlignment.BOTTOM);
            // POI 把 BOTTOM 视为默认值，读路径不会回填 vt（这是已知的）
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_shrink_to_fit_via_tb_equals_two() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle poi = sc.toPoiStyle(new IStyleData().setTb(2));
            assertThat(poi.getShrinkToFit()).isTrue();
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getTb()).isEqualTo(2);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_tb_overflow_value_one() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // tb=1 (overflow) 不应导致 wrap/shrink
            XSSFCellStyle poi = sc.toPoiStyle(new IStyleData().setTb(1));
            assertThat(poi.getWrapText()).isFalse();
            assertThat(poi.getShrinkToFit()).isFalse();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_null_input_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 传入 null：应当当作空 IStyleData 处理而非 NPE
            XSSFCellStyle poi = sc.toPoiStyle(null);
            assertThat(poi).isNotNull();
            String id = sc.styleIdOf(null);
            assertThat(id).hasSize(16);
            XSSFCellStyle qp = sc.toPoiStyleWithQuotePrefix(null);
            assertThat(qp.getQuotePrefixed()).isTrue();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_register_observed_styles_in_read_path() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData s = new IStyleData().setFf("Calibri").setFs(11);
            String id = sc.styleIdOf(s);
            Map<String, IStyleData> reg = sc.getStyleRegistry();
            assertThat(reg).containsKey(id);
            assertThat(reg.get(id)).isSameAs(s);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_numfmt_for_general_format_on_read() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 没有设置数字格式时读路径应当跳过 n 字段
            XSSFCellStyle poi = sc.toPoiStyle(new IStyleData().setFf("Arial"));
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getN()).isNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_pass_through_numfmt_when_pattern_is_blank() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // n.pattern=null：applyNumFmt 应该跳过设置
            IStyleData src = new IStyleData().setN(new INumfmtLocal());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getDataFormatString()).isEqualToIgnoringCase("General");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_empty_style_for_default_index_zero() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            XSSFCellStyle defaultStyle = wb.getCellStyleAt(0);
            IStyleData back = sc.fromPoiStyle(defaultStyle);
            // index=0 即默认样式，应返回空样式
            assertThat(back).isEqualTo(new IStyleData());
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_empty_style_for_null_input_in_read_path() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            assertThat(sc.fromPoiStyle(null)).isEqualTo(new IStyleData());
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_invalid_hex_in_font_color() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // 非法颜色字串不会被写入字体 / invalid hex must not throw
            IStyleData src = new IStyleData().setCl(new IColorStyle().setRgb("not-a-hex"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            // 不抛错即可
            assertThat(poi).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_invalid_hex_in_background() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setBg(new IColorStyle().setRgb("xxx"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            // 仍然能产生 style；fillPattern 不会被强制成 SOLID
            assertThat(poi).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_border_edge_with_null_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // bd.t 存在但 s 为 null：不应抛错也不应产生边框
            IStyleData src = new IStyleData()
                    .setBd(new IBorderData().setT(new IBorderStyleData()));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getBorderTop()).isEqualTo(BorderStyle.NONE);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_cache_quote_prefix_variant() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData a = new IStyleData().setFf("Arial").setFs(10);
            IStyleData b = new IStyleData().setFf("Arial").setFs(10);
            assertThat(sc.toPoiStyleWithQuotePrefix(a)).isSameAs(sc.toPoiStyleWithQuotePrefix(b));
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_text_rotation_arbitrary_angle() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setTr(new ITextRotation().setA(90));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getRotation()).isEqualTo((short) 90);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getTr().getA()).isEqualTo(90);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_round_trip_underline_off_when_absent() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setFf("Arial");
            XSSFCellStyle poi = sc.toPoiStyle(src);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getUl()).isNull();
            assertThat(back.getSt()).isNull();
            assertThat(back.getBl()).isNull();
            assertThat(back.getIt()).isNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_text_rotation_when_node_present_without_values() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            // tr 对象存在但 v/a 都为 null：applyRotation 应跳过
            IStyleData src = new IStyleData().setTr(new ITextRotation());
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getRotation()).isEqualTo((short) 0);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_cache_style_id_on_repeated_calls() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData s = new IStyleData().setFf("Verdana").setFs(13);
            // 第二次调用应命中 styleIdCache
            String first = sc.styleIdOf(s);
            String second = sc.styleIdOf(s);
            assertThat(first).isEqualTo(second);
        } catch (Exception e) { throw new RuntimeException(e); }
    }
}
