/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */
package io.github.autoffice.univer.converter;

import io.github.autoffice.univer.model.BooleanNumber;
import io.github.autoffice.univer.model.IColorStyle;
import io.github.autoffice.univer.model.IDocumentData;
import io.github.autoffice.univer.model.IStyleData;
import io.github.autoffice.univer.model.ITextDecoration;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * RichTextConverter partial 分支精确补全：
 * - null/空 dataStream
 * - run 缺 st/ed
 * - ts.ul/st/cl 各种 null 边界
 * - customRanges 的非法 rangeType / props
 * - fontToStyle 各 if 的另一边
 */
class RichTextConverterPartialBranchTest {

    @Test
    void should_return_empty_rts_for_null_input() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            XSSFRichTextString rts = rc.toPoi(null);
            assertThat(rts.getString()).isEmpty();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_return_empty_rts_for_null_body() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData p = new IDocumentData(); // body=null
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts.getString()).isEmpty();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_runs_with_null_st_or_ed() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            List<IDocumentData.TextRun> runs = new ArrayList<>();
            runs.add(null); // null run
            runs.add(new IDocumentData.TextRun().setEd(5)); // st=null
            runs.add(new IDocumentData.TextRun().setSt(0)); // ed=null
            runs.add(new IDocumentData.TextRun().setSt(0).setEd(5)
                    .setTs(new IStyleData().setFf("Arial")));
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                    .setDataStream("hello\r\n")
                    .setTextRuns(runs));
            XSSFRichTextString rts = rc.toPoi(p);
            // 不抛错；至少 1 个有效 run
            assertThat(rts.numFormattingRuns()).isGreaterThanOrEqualTo(1);
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_run_when_st_equals_ed() {
        // 覆盖 L96: st >= ed → continue
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                    .setDataStream("hello\r\n")
                    .setTextRuns(Collections.singletonList(
                            new IDocumentData.TextRun().setSt(3).setEd(3))));
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts.getString()).isEqualTo("hello");
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_run_with_ul_object_null_s() {
        // 覆盖 L191: ts.ul != null 但 s == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IStyleData ts = new IStyleData()
                    .setUl(new ITextDecoration()) // ul 存在但 s=null
                    .setSt(new ITextDecoration()); // st 存在但 s=null
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                    .setDataStream("text\r\n")
                    .setTextRuns(Collections.singletonList(
                            new IDocumentData.TextRun().setSt(0).setEd(4).setTs(ts))));
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_run_with_cl_null_rgb() {
        // 覆盖 L197: ts.cl != null 但 rgb == null
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IStyleData ts = new IStyleData().setCl(new IColorStyle()); // rgb=null
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                    .setDataStream("text\r\n")
                    .setTextRuns(Collections.singletonList(
                            new IDocumentData.TextRun().setSt(0).setEd(4).setTs(ts))));
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_handle_run_with_invalid_color_rgb() {
        // 覆盖 L199: argb == null（rgb 是非法 hex）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IStyleData ts = new IStyleData().setCl(new IColorStyle().setRgb("invalid"));
            IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                    .setDataStream("text\r\n")
                    .setTextRuns(Collections.singletonList(
                            new IDocumentData.TextRun().setSt(0).setEd(4).setTs(ts))));
            XSSFRichTextString rts = rc.toPoi(p);
            assertThat(rts).isNotNull();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_read_from_null_rts() {
        // 覆盖 L134: rts == null 分支
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            IDocumentData back = rc.fromPoi(null);
            assertThat(back).isNotNull();
            // 应至少有一个段落
            assertThat(back.getBody().getParagraphs()).isNotEmpty();
        } catch (Exception e) { throw new RuntimeException(e); }
    }

    @Test
    void should_skip_when_first_hyperlink_url_doc_null() {
        // 覆盖 L253: p == null 或 p.body == null
        assertThat(RichTextConverter.firstHyperlinkUrl(null)).isNull();
        IDocumentData noBody = new IDocumentData(); // body=null
        assertThat(RichTextConverter.firstHyperlinkUrl(noBody)).isNull();
    }

    @Test
    void should_skip_custom_range_with_non_hyperlink_type() {
        // 覆盖 L270: rangeType 不是 hyperlink → continue
        IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("text\r\n"));
        List<Map<String, Object>> ranges = new ArrayList<>();
        Map<String, Object> r = new LinkedHashMap<>();
        r.put("rangeType", 99); // 非 hyperlink
        Map<String, Object> props = new LinkedHashMap<>();
        props.put("url", "https://x.com");
        r.put("properties", props);
        ranges.add(r);
        p.getBody().putExtra("customRanges", ranges);
        assertThat(RichTextConverter.firstHyperlinkUrl(p)).isNull();
    }

    @Test
    void should_skip_custom_range_with_no_props() {
        // 覆盖 L274: props 不是 Map
        IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("text\r\n"));
        List<Map<String, Object>> ranges = new ArrayList<>();
        Map<String, Object> r = new LinkedHashMap<>();
        r.put("rangeType", 0);
        // 不设 properties
        ranges.add(r);
        p.getBody().putExtra("customRanges", ranges);
        assertThat(RichTextConverter.firstHyperlinkUrl(p)).isNull();
    }

    @Test
    void should_skip_custom_range_with_empty_url() {
        // 覆盖 L276: url 为空串或非 String
        IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("text\r\n"));
        List<Map<String, Object>> ranges = new ArrayList<>();
        Map<String, Object> r = new LinkedHashMap<>();
        r.put("rangeType", 0);
        Map<String, Object> props = new LinkedHashMap<>();
        props.put("url", ""); // 空串
        r.put("properties", props);
        ranges.add(r);
        p.getBody().putExtra("customRanges", ranges);
        assertThat(RichTextConverter.firstHyperlinkUrl(p)).isNull();
    }

    @Test
    void should_skip_attach_hyperlink_with_null_inputs() {
        // 覆盖 L292: p == null / body == null / url == null / url.isEmpty
        RichTextConverter.attachHyperlink(null, "u", 0, 1); // null p
        IDocumentData p1 = new IDocumentData(); // body=null
        RichTextConverter.attachHyperlink(p1, "u", 0, 1);
        IDocumentData p2 = new IDocumentData().setBody(new IDocumentData.Body());
        RichTextConverter.attachHyperlink(p2, null, 0, 1); // null url
        RichTextConverter.attachHyperlink(p2, "", 0, 1); // 空 url
        // 不抛错即可
        assertThat(p2.getBody().getExtras()).doesNotContainKey("customRanges");
    }

    @Test
    void should_create_customRanges_list_if_absent_in_attachHyperlink() {
        // 覆盖 L298 的 else 分支（raw 不是 List → 新建）
        IDocumentData p = new IDocumentData().setBody(new IDocumentData.Body()
                .setDataStream("text\r\n"));
        // 不预设 extras
        RichTextConverter.attachHyperlink(p, "https://x.com", 0, 4);
        assertThat(p.getBody().getExtras()).containsKey("customRanges");
    }

    @Test
    void should_handle_fontToStyle_default_branches() throws Exception {
        // 覆盖 fontToStyle 各 if 的另一侧（font 名称为 null，font 大小为 0 等）
        // 通过空 XSSFRichTextString 让 numFormattingRuns=0 来跳过 run 处理
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            RichTextConverter rc = new RichTextConverter(wb);
            XSSFRichTextString rts = new XSSFRichTextString("text");
            // 创建一个 font 没设任何属性
            org.apache.poi.xssf.usermodel.XSSFFont font = wb.createFont();
            font.setBold(false);
            rts.applyFont(0, 4, font);
            IDocumentData back = rc.fromPoi(rts);
            assertThat(back).isNotNull();
        }
    }
}
