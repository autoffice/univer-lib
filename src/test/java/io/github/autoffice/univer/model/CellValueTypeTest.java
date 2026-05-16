/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */
package io.github.autoffice.univer.model;

import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * 覆盖 CellValueType 的 getValue / of(int) 全部分支 + 未知值异常路径。
 */
class CellValueTypeTest {

    @Test
    void should_map_numeric_codes_to_enum_constants() {
        assertThat(CellValueType.of(1)).isEqualTo(CellValueType.STRING);
        assertThat(CellValueType.of(2)).isEqualTo(CellValueType.NUMBER);
        assertThat(CellValueType.of(3)).isEqualTo(CellValueType.BOOLEAN);
        assertThat(CellValueType.of(4)).isEqualTo(CellValueType.FORCE_TEXT);
    }

    @Test
    void should_expose_numeric_value_for_each_constant() {
        assertThat(CellValueType.STRING.getValue()).isEqualTo(1);
        assertThat(CellValueType.NUMBER.getValue()).isEqualTo(2);
        assertThat(CellValueType.BOOLEAN.getValue()).isEqualTo(3);
        assertThat(CellValueType.FORCE_TEXT.getValue()).isEqualTo(4);
    }

    @Test
    void should_throw_for_unknown_numeric_code() {
        assertThatThrownBy(() -> CellValueType.of(99))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("Unknown CellValueType");
    }

    @Test
    void should_serialize_as_number_via_jackson() throws Exception {
        String json = JsonMapper.get().writeValueAsString(CellValueType.NUMBER);
        assertThat(json).isEqualTo("2");
        CellValueType back = JsonMapper.get().readValue("4", CellValueType.class);
        assertThat(back).isEqualTo(CellValueType.FORCE_TEXT);
    }
}
