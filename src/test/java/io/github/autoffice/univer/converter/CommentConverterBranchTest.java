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

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.autoffice.univer.util.JsonMapper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * CommentConverter 边角分支：null inputs、坏 key、坏值、null sheet 等。
 */
class CommentConverterBranchTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_return_empty_for_null_or_no_comments_sheet() throws Exception {
        ObjectNode out = CommentConverter.readSheetComments(null, mapper);
        assertThat(out.size()).isEqualTo(0);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            // 没注释的 sheet 也应该返回空
            ObjectNode out2 = CommentConverter.readSheetComments(sh, mapper);
            assertThat(out2.size()).isEqualTo(0);
        }
    }

    @Test
    void should_skip_write_with_invalid_inputs() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            CommentConverter.writeSheetComments(null, mapper.createObjectNode());
            CommentConverter.writeSheetComments(sh, null);
            CommentConverter.writeSheetComments(sh, mapper.createArrayNode()); // 非 object
            CommentConverter.writeSheetComments(sh, mapper.createObjectNode()); // 空 object
            assertThat(sh.hasComments()).isFalse();
        }
    }

    @Test
    void should_skip_invalid_row_or_col_keys() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode rowMap = mapper.createObjectNode();
            // 非数字 row → 跳过
            ObjectNode validCol = mapper.createObjectNode();
            ObjectNode note = mapper.createObjectNode();
            note.put("note", "should not appear");
            validCol.set("0", note);
            rowMap.set("not-a-number", validCol);
            // 空字符串 row
            rowMap.set("", validCol);

            // 数字 row 但 colMap 是数组 → 跳过
            rowMap.set("0", mapper.createArrayNode());

            // 数字 row 且 colMap 是 object，但 col 非数字
            ObjectNode r1 = mapper.createObjectNode();
            r1.set("not-a-col", note);
            // 空 col
            r1.set("", note);
            // 数字 col 但 noteNode 非 object
            r1.set("3", mapper.createArrayNode());
            rowMap.set("1", r1);

            CommentConverter.writeSheetComments(sh, rowMap);
            // 没有有效注释被写入
            assertThat(sh.hasComments()).isFalse();
        }
    }

    @Test
    void should_write_valid_note_when_keys_valid() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sh = wb.createSheet("S");
            ObjectNode rowMap = mapper.createObjectNode();
            ObjectNode r0 = mapper.createObjectNode();
            ObjectNode note = mapper.createObjectNode();
            note.put("note", "hi");
            note.put("show", true);
            r0.set("0", note);
            rowMap.set("0", r0);
            CommentConverter.writeSheetComments(sh, rowMap);
            assertThat(sh.hasComments()).isTrue();
        }
    }
}
