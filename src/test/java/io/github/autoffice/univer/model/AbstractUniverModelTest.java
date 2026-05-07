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
