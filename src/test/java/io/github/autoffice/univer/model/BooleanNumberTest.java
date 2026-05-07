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
