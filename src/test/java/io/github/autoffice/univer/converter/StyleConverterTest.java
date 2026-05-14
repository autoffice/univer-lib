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

    /**
     * 回归：当 cellXf 在 XML 中未显式声明 applyBorder 属性时，POI 高层 API 会把
     * borderTop/Left/Right/Bottom 全部返回 NONE，导致边框丢失（test2.xlsx 中
     * "丰富的单元格表现" sheet 的 E3 dotted red 边框就是这种情况）。修复后应该
     * 从 CTBorder 兜底读取，让边框正确进入 IStyleData。
     */
    @Test
    void should_read_border_when_apply_border_flag_is_unset() throws Exception {
        try (java.io.FileInputStream fis = new java.io.FileInputStream("src/test/resources/test2.xlsx");
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheet("丰富的单元格表现");
            assertThat(sheet).as("fixture 必须包含目标 sheet").isNotNull();
            XSSFCell e3 = sheet.getRow(2).getCell(4);
            assertThat(e3).as("E3 应存在").isNotNull();
            XSSFCellStyle cs = e3.getCellStyle();

            // 触发 POI 的 bug：高层 API 因 applyBorder 未显式设置而返回 NONE
            assertThat(cs.getBorderTop()).isEqualTo(org.apache.poi.ss.usermodel.BorderStyle.NONE);

            // 但 StyleConverter 应该兜底读到 dotted red 边框（Univer s=3, color #ff0000）
            StyleConverter sc = new StyleConverter(wb);
            IStyleData out = sc.fromPoiStyle(cs);
            assertThat(out.getBd()).as("应通过 CTBorder 兜底读到边框").isNotNull();
            assertThat(out.getBd().getT()).isNotNull();
            assertThat(out.getBd().getT().getS()).as("DOTTED").isEqualTo(3);
            assertThat(out.getBd().getT().getCl().getRgb().toLowerCase()).isEqualTo("#ff0000");
            assertThat(out.getBd().getB().getS()).isEqualTo(3);
            assertThat(out.getBd().getL().getS()).isEqualTo(3);
            assertThat(out.getBd().getR().getS()).isEqualTo(3);
        }
    }
}
