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

import com.fasterxml.jackson.databind.ObjectMapper;
import io.github.autoffice.univer.util.JsonMapper;
import org.junit.jupiter.api.Test;
import static org.assertj.core.api.Assertions.assertThat;

class BooleanNumberTest {
    @Test
    void should_serialize_as_number_and_deserialize_back() throws Exception {
        ObjectMapper m = JsonMapper.get();
        assertThat(m.writeValueAsString(BooleanNumber.TRUE)).isEqualTo("1");
        assertThat(m.writeValueAsString(BooleanNumber.FALSE)).isEqualTo("0");
        assertThat(m.readValue("1", BooleanNumber.class)).isEqualTo(BooleanNumber.TRUE);
        assertThat(m.readValue("0", BooleanNumber.class)).isEqualTo(BooleanNumber.FALSE);
    }
}
