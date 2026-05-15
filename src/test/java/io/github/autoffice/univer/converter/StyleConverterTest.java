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

import io.github.autoffice.univer.model.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class StyleConverterTest {

    @Test
    void should_roundtrip_font_and_alignment() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setFf("Arial").setFs(12)
                .setBl(BooleanNumber.TRUE).setIt(BooleanNumber.TRUE)
                .setHt(1).setVt(2).setCl(new IColorStyle().setRgb("#ff0000"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getAlignment()).isEqualTo(HorizontalAlignment.LEFT);
            assertThat(poi.getVerticalAlignment()).isEqualTo(VerticalAlignment.CENTER);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getFf()).isEqualTo("Arial");
            assertThat(back.getFs()).isEqualTo(12);
            assertThat(back.getBl()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getIt()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getHt()).isEqualTo(1);
            assertThat(back.getVt()).isEqualTo(2);
            assertThat(back.getCl().getRgb()).isEqualToIgnoringCase("#ff0000");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_roundtrip_background_border_wrap_rotation() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData()
                .setBg(new IColorStyle().setRgb("#00ff00"))
                .setBd(new IBorderData().setT(new IBorderStyleData().setS(1)
                        .setCl(new IColorStyle().setRgb("#000000"))))
                .setTb(3)
                .setTr(new ITextRotation().setA(45));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getBg().getRgb()).isEqualToIgnoringCase("#00ff00");
            assertThat(back.getBd().getT().getS()).isEqualTo(1);
            assertThat(back.getTb()).isEqualTo(3);
            assertThat(back.getTr().getA()).isEqualTo(45);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_roundtrip_number_format() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setN(new INumfmtLocal().setPattern("yyyy-mm-dd"));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getDataFormatString()).isEqualTo("yyyy-mm-dd");
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getN().getPattern()).isEqualTo("yyyy-mm-dd");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_cache_equal_styles_to_one_poi_style() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData a = new IStyleData().setFf("Arial").setFs(10);
            IStyleData b = new IStyleData().setFf("Arial").setFs(10);
            assertThat(sc.toPoiStyle(a)).isSameAs(sc.toPoiStyle(b));
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_generate_stable_style_id() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData a = new IStyleData().setFf("Arial").setFs(10);
            IStyleData b = new IStyleData().setFf("Arial").setFs(10);
            assertThat(sc.styleIdOf(a)).isEqualTo(sc.styleIdOf(b));
            assertThat(sc.styleIdOf(a)).hasSize(16);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_vertical_text_rotation() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData().setTr(new ITextRotation().setV(BooleanNumber.TRUE));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            assertThat(poi.getRotation()).isEqualTo((short)255);
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getTr().getV()).isEqualTo(BooleanNumber.TRUE);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_underline_and_strikethrough() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            StyleConverter sc = new StyleConverter(wb);
            IStyleData src = new IStyleData()
                .setUl(new ITextDecoration().setS(BooleanNumber.TRUE))
                .setSt(new ITextDecoration().setS(BooleanNumber.TRUE));
            XSSFCellStyle poi = sc.toPoiStyle(src);
            XSSFFont font = wb.getFontAt(poi.getFontIndex());
            assertThat(font.getUnderline()).isNotEqualTo((byte)0);
            assertThat(font.getStrikeout()).isTrue();
            IStyleData back = sc.fromPoiStyle(poi);
            assertThat(back.getUl().getS()).isEqualTo(BooleanNumber.TRUE);
            assertThat(back.getSt().getS()).isEqualTo(BooleanNumber.TRUE);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

}
