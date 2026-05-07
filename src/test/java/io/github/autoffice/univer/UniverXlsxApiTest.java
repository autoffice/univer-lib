package io.github.autoffice.univer;

import io.github.autoffice.univer.model.IWorkbookData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class UniverXlsxApiTest {

    @Test
    void should_expose_read_facade_and_throw_not_implemented() {
        assertThatThrownBy(() -> UniverXlsx.read(new ByteArrayInputStream(new byte[0])))
                .isInstanceOf(UniverXlsxReadException.class);
    }

    @Test
    void should_expose_write_facade_and_throw_not_implemented() {
        assertThatThrownBy(() -> UniverXlsx.write(new IWorkbookData(), new ByteArrayOutputStream()))
                .isInstanceOf(UniverXlsxWriteException.class);
    }

    @Test
    void should_build_options_with_defaults() {
        UniverXlsxOptions opts = UniverXlsxOptions.defaults();
        assertThat(opts.isStrictMode()).isFalse();
        assertThat(opts.isWriteSidecar()).isTrue();
        assertThat(opts.isPrettyJson()).isFalse();
        assertThat(opts.getLocale()).isEqualTo("enUS");
    }

    @Test
    void should_build_options_with_custom_values() {
        UniverXlsxOptions opts = UniverXlsxOptions.builder()
                .strictMode(true)
                .writeSidecar(false)
                .prettyJson(true)
                .locale("zhCN")
                .build();
        assertThat(opts.isStrictMode()).isTrue();
        assertThat(opts.isWriteSidecar()).isFalse();
        assertThat(opts.isPrettyJson()).isTrue();
        assertThat(opts.getLocale()).isEqualTo("zhCN");
    }
}
