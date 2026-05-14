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
package io.github.autoffice.univer.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;
import io.github.autoffice.univer.util.JsonMapper;

class AbstractUniverModelTest {
    @Data @EqualsAndHashCode(callSuper = true)
    @JsonInclude(JsonInclude.Include.NON_NULL)
    static class Sample extends AbstractUniverModel { private String name; }

    @Test
    void should_roundtrip_unknown_fields_via_extras() throws Exception {
        Sample s = JsonMapper.get().readValue("{\"name\":\"a\",\"future\":42}", Sample.class);
        assertThat(s.getName()).isEqualTo("a");
        assertThat(s.getExtras()).containsEntry("future", 42);
        String json = JsonMapper.get().writeValueAsString(s);
        assertThat(json).contains("\"future\":42");
    }
}
