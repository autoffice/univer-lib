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
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 通过反射覆盖 ConditionalFormattingConverter 的私有 extractHex / thresholdToValue 等剩余分支。
 */
class ConditionalFormattingConverterPrivateTest {

    private final ObjectMapper mapper = JsonMapper.get();

    @Test
    void should_handle_extractHex_for_null_xssfColor_extendedColor() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "extractHex", org.apache.poi.ss.usermodel.Color.class);
        m.setAccessible(true);

        // null → null
        assertThat(m.invoke(null, (Object) null)).isNull();

        // XSSFColor → hex
        XSSFColor xssf = new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xAA, (byte) 0xBB, (byte) 0xCC}, null);
        assertThat((String) m.invoke(null, xssf)).isEqualToIgnoringCase("#aabbcc");

        // 不属于 XSSFColor / ExtendedColor 的实现 → null
        org.apache.poi.ss.usermodel.Color other = new org.apache.poi.ss.usermodel.Color() {};
        assertThat(m.invoke(null, other)).isNull();
    }

    @Test
    void should_threshold_to_value_handles_null_input() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "thresholdToValue",
                ConditionalFormattingThreshold.class, ObjectMapper.class);
        m.setAccessible(true);

        // null threshold → 默认 num/0
        ObjectNode out = (ObjectNode) m.invoke(null, null, mapper);
        assertThat(out.path("type").asText()).isEqualTo("num");
        assertThat(out.path("value").asInt()).isEqualTo(0);
    }

    @Test
    void should_map_each_range_type_via_reflection() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "mapRangeType", ConditionalFormattingThreshold.RangeType.class);
        m.setAccessible(true);
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.MIN)).isEqualTo("min");
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.MAX)).isEqualTo("max");
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.PERCENT)).isEqualTo("percent");
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.PERCENTILE)).isEqualTo("percentile");
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.FORMULA)).isEqualTo("formula");
        assertThat((String) m.invoke(null, ConditionalFormattingThreshold.RangeType.NUMBER)).isEqualTo("num");
    }

    @Test
    void should_map_each_comparison_operator_via_reflection() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "mapComparisonOperator", byte.class);
        m.setAccessible(true);
        // ComparisonOperator: NO_COMPARISON=0, BETWEEN=1, NOT_BETWEEN=2, EQUAL=3, NOT_EQUAL=4,
        //                     GT=5, LT=6, GE=7, LE=8
        assertThat((String) m.invoke(null, (byte) 3)).isEqualTo("equal");
        assertThat((String) m.invoke(null, (byte) 4)).isEqualTo("notEqual");
        assertThat((String) m.invoke(null, (byte) 5)).isEqualTo("greaterThan");
        assertThat((String) m.invoke(null, (byte) 6)).isEqualTo("lessThan");
        assertThat((String) m.invoke(null, (byte) 7)).isEqualTo("greaterThanOrEqual");
        assertThat((String) m.invoke(null, (byte) 8)).isEqualTo("lessThanOrEqual");
        assertThat((String) m.invoke(null, (byte) 1)).isEqualTo("between");
        assertThat((String) m.invoke(null, (byte) 2)).isEqualTo("notBetween");
        // 未知 → equal
        assertThat((String) m.invoke(null, (byte) 99)).isEqualTo("equal");
    }

    @Test
    void should_map_univer_operator_to_poi_via_reflection() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "univerOperatorToPoi", String.class);
        m.setAccessible(true);
        // null
        assertThat((Byte) m.invoke(null, (Object) null)).isEqualTo((byte) 3); // EQUAL
        // unknown
        assertThat((Byte) m.invoke(null, "unknownOp")).isEqualTo((byte) 3);
        // 各 case
        String[] ops = {"equal", "notEqual", "greaterThan", "lessThan",
                        "greaterThanOrEqual", "lessThanOrEqual", "between", "notBetween"};
        for (String op : ops) {
            Object code = m.invoke(null, op);
            assertThat((Byte) code).as("op=%s", op).isNotNull();
        }
    }

    @Test
    void should_strip_leading_eq_via_reflection() throws Exception {
        Method m = ConditionalFormattingConverter.class.getDeclaredMethod(
                "stripLeadingEq", String.class);
        m.setAccessible(true);
        assertThat((String) m.invoke(null, "=A1")).isEqualTo("A1");
        assertThat((String) m.invoke(null, "no-eq")).isEqualTo("no-eq");
        assertThat(m.invoke(null, (Object) null)).isNull();
    }
}
