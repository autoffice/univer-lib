/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */
package io.github.autoffice.univer.util;

import com.fasterxml.jackson.core.type.TypeReference;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class IntegerKeyDeserializerTest {

    @Test
    void should_deserialize_integer_keys_in_nested_map() throws Exception {
        String json = "{\"1\":{\"2\":\"v\"},\"3\":{\"4\":\"w\"}}";
        Map<Integer, Map<Integer, String>> out =
                JsonMapper.get().readValue(json, new TypeReference<Map<Integer, Map<Integer, String>>>() {});
        assertThat(out).containsKey(1).containsKey(3);
        assertThat(out.get(1)).containsEntry(2, "v");
        assertThat(out.get(3)).containsEntry(4, "w");
    }
}
