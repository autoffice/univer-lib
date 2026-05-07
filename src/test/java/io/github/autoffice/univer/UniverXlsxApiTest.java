package io.github.autoffice.univer;

import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Collections;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class UniverXlsxApiTest {

    @Test
    void should_throw_read_exception_on_invalid_input() {
        assertThatThrownBy(() -> UniverXlsx.read(new ByteArrayInputStream(new byte[]{1, 2, 3})))
                .isInstanceOf(UniverXlsxReadException.class);
    }

    @Test
    void should_write_empty_workbook_without_error() throws Exception {
        IWorkbookData empty = new IWorkbookData();
        empty.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        empty.setSheetOrder(Collections.singletonList("s1"));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(empty, out);
        assertThat(out.size()).isGreaterThan(0);
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
