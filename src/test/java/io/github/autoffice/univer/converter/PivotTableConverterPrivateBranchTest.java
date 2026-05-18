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

import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;

import java.lang.reflect.Method;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 通过反射覆盖 PivotTableConverter 的私有 helper 方法的边角分支：
 * - subtotalToUniverName: null + 各种 enum
 * - subtotalToPoi: null/empty + 各种字符串
 */
class PivotTableConverterPrivateBranchTest {

    @Test
    void should_map_each_subtotal_enum_to_univer_name() throws Exception {
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "subtotalToUniverName", STDataConsolidateFunction.Enum.class);
        m.setAccessible(true);

        // null → "sum" (default)
        assertThat((String) m.invoke(null, (Object) null)).isEqualTo("sum");
        // 主要 enum 值（POI 5.x 中 STDataConsolidateFunction 常量名）
        assertThat((String) m.invoke(null, STDataConsolidateFunction.SUM)).isEqualTo("sum");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.COUNT)).isEqualTo("count");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.AVERAGE)).isEqualTo("average");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.MAX)).isEqualTo("max");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.MIN)).isEqualTo("min");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.PRODUCT)).isEqualTo("product");
        // countNums → "countNumbers"
        assertThat((String) m.invoke(null, STDataConsolidateFunction.COUNT_NUMS))
                .isEqualTo("countNumbers");
        assertThat((String) m.invoke(null, STDataConsolidateFunction.VAR)).isEqualTo("var");
    }

    @Test
    void should_map_each_subtotal_string_to_poi() throws Exception {
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "subtotalToPoi", String.class);
        m.setAccessible(true);

        // null/empty → SUM
        assertThat(m.invoke(null, (Object) null)).isEqualTo(DataConsolidateFunction.SUM);
        assertThat(m.invoke(null, "")).isEqualTo(DataConsolidateFunction.SUM);

        // 各 case
        assertThat(m.invoke(null, "sum")).isEqualTo(DataConsolidateFunction.SUM);
        assertThat(m.invoke(null, "count")).isEqualTo(DataConsolidateFunction.COUNT);
        assertThat(m.invoke(null, "countNumbers")).isEqualTo(DataConsolidateFunction.COUNT_NUMS);
        assertThat(m.invoke(null, "average")).isEqualTo(DataConsolidateFunction.AVERAGE);
        assertThat(m.invoke(null, "max")).isEqualTo(DataConsolidateFunction.MAX);
        assertThat(m.invoke(null, "min")).isEqualTo(DataConsolidateFunction.MIN);
        assertThat(m.invoke(null, "product")).isEqualTo(DataConsolidateFunction.PRODUCT);
        assertThat(m.invoke(null, "stdDev")).isEqualTo(DataConsolidateFunction.STD_DEV);
        assertThat(m.invoke(null, "stdDevp")).isEqualTo(DataConsolidateFunction.STD_DEVP);
        assertThat(m.invoke(null, "var")).isEqualTo(DataConsolidateFunction.VAR);
        assertThat(m.invoke(null, "varp")).isEqualTo(DataConsolidateFunction.VARP);
        // unknown → SUM
        assertThat(m.invoke(null, "weird-unknown")).isEqualTo(DataConsolidateFunction.SUM);
    }

    @Test
    void should_build_fieldId_with_area_index_combo() throws Exception {
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "buildFieldId", String.class, int.class, int.class);
        m.setAccessible(true);
        // 覆盖此 helper 方法
        assertThat((String) m.invoke(null, "row", 0, 0)).isEqualTo("row-0-0");
        assertThat((String) m.invoke(null, "column", 1, 2)).isEqualTo("column-1-2");
        assertThat((String) m.invoke(null, "filter", 5, 0)).isEqualTo("filter-5-0");
        assertThat((String) m.invoke(null, "value", 2, 1)).isEqualTo("value-2-1");
    }

    @Test
    void should_emptyToNull_handle_inputs() throws Exception {
        Method m = PivotTableConverter.class.getDeclaredMethod("emptyToNull", String.class);
        m.setAccessible(true);
        assertThat(m.invoke(null, (Object) null)).isNull();
        assertThat(m.invoke(null, "")).isNull();
        assertThat(m.invoke(null, "x")).isEqualTo("x");
    }

    @Test
    void should_handle_collectUnsupportedFeatures_with_null_def() throws Exception {
        // 覆盖 L248: def == null
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "collectUnsupportedFeatures",
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition.class,
                com.fasterxml.jackson.databind.node.ArrayNode.class,
                com.fasterxml.jackson.databind.node.ArrayNode.class);
        m.setAccessible(true);

        com.fasterxml.jackson.databind.ObjectMapper mapper = io.github.autoffice.univer.util.JsonMapper.get();
        com.fasterxml.jackson.databind.node.ArrayNode unsupported = mapper.createArrayNode();
        com.fasterxml.jackson.databind.node.ArrayNode detected = mapper.createArrayNode();
        m.invoke(null, null, unsupported, detected);
        // unsupported 应包含基础特性
        assertThat(unsupported.size()).isGreaterThan(0);
        assertThat(detected.size()).isEqualTo(0);
    }

    @Test
    void should_handle_resolveFieldNames_with_null_cacheDef() throws Exception {
        // 覆盖 L264: cacheDefinition == null
        // XSSFPivotTable 是 final class，无法用 Proxy mock
        // 跳过此测试，依赖 round-trip 测试覆盖
    }

    @Test
    void should_handle_resolvePivotFieldMap_with_null_pivotFields() throws Exception {
        // 覆盖 L281: pivotFields == null
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "resolvePivotFieldMap",
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition.class);
        m.setAccessible(true);

        // 创建一个 def 不带 pivotFields
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition def =
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition.Factory.newInstance();
        // 默认无 pivotFields
        Object result = m.invoke(null, def);
        assertThat(result).isInstanceOf(java.util.Map.class);
        assertThat(((java.util.Map<?,?>) result)).isEmpty();
    }

    @Test
    void should_handle_addAxisFields_with_null_rowFields() throws Exception {
        // 覆盖 L296: rowFields == null
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "addAxisFields",
                com.fasterxml.jackson.databind.node.ArrayNode.class,
                String.class,
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowFields.class,
                java.util.Map.class,
                java.util.Map.class,
                com.fasterxml.jackson.databind.ObjectMapper.class);
        m.setAccessible(true);

        com.fasterxml.jackson.databind.ObjectMapper mapper = io.github.autoffice.univer.util.JsonMapper.get();
        com.fasterxml.jackson.databind.node.ArrayNode out = mapper.createArrayNode();
        m.invoke(null, out, "row", null, new java.util.LinkedHashMap<>(),
                new java.util.LinkedHashMap<>(), mapper);
        assertThat(out.size()).isEqualTo(0);
    }

    @Test
    void should_handle_addPageFields_with_null() throws Exception {
        // 覆盖 addPageFields null 分支
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "addPageFields",
                com.fasterxml.jackson.databind.node.ArrayNode.class,
                java.util.Map.class,
                java.util.Map.class,
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageFields.class,
                com.fasterxml.jackson.databind.ObjectMapper.class);
        m.setAccessible(true);

        com.fasterxml.jackson.databind.ObjectMapper mapper = io.github.autoffice.univer.util.JsonMapper.get();
        com.fasterxml.jackson.databind.node.ArrayNode out = mapper.createArrayNode();
        m.invoke(null, out, new java.util.LinkedHashMap<>(),
                new java.util.LinkedHashMap<>(), null, mapper);
        assertThat(out.size()).isEqualTo(0);
    }

    @Test
    void should_handle_addValueFields_with_null() throws Exception {
        // 覆盖 addValueFields null 分支
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "addValueFields",
                com.fasterxml.jackson.databind.node.ArrayNode.class,
                java.util.Map.class,
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields.class,
                com.fasterxml.jackson.databind.ObjectMapper.class);
        m.setAccessible(true);

        com.fasterxml.jackson.databind.ObjectMapper mapper = io.github.autoffice.univer.util.JsonMapper.get();
        com.fasterxml.jackson.databind.node.ArrayNode out = mapper.createArrayNode();
        m.invoke(null, out, new java.util.LinkedHashMap<>(), null, mapper);
        assertThat(out.size()).isEqualTo(0);
    }

    @Test
    void should_handle_toAreaReference_with_invalid_indices() throws Exception {
        // 覆盖 L594: 各种 -1 索引返回 null
        Method m = PivotTableConverter.class.getDeclaredMethod(
                "toAreaReference", com.fasterxml.jackson.databind.JsonNode.class);
        m.setAccessible(true);

        com.fasterxml.jackson.databind.ObjectMapper mapper = io.github.autoffice.univer.util.JsonMapper.get();

        // 缺字段
        ObjectNode empty = mapper.createObjectNode();
        assertThat(m.invoke(null, empty)).isNull();

        // 负数 startRow
        ObjectNode bad1 = mapper.createObjectNode();
        bad1.put("startRow", -1);
        bad1.put("endRow", 5);
        bad1.put("startColumn", 0);
        bad1.put("endColumn", 5);
        assertThat(m.invoke(null, bad1)).isNull();

        // 有效 range
        ObjectNode valid = mapper.createObjectNode();
        valid.put("startRow", 0);
        valid.put("endRow", 5);
        valid.put("startColumn", 0);
        valid.put("endColumn", 5);
        assertThat(m.invoke(null, valid)).isNotNull();
    }
}
