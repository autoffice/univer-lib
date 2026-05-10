package io.github.autoffice.univer;

import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
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
    void should_roundtrip_data_validation_resources_through_public_api() throws Exception {
        IWorkbookData src = new IWorkbookData();
        src.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        src.setSheetOrder(Collections.singletonList("s1"));
        List<Object> resources = new ArrayList<>();
        Map<String, Object> entry = new LinkedHashMap<>();
        entry.put("name", "SHEET_DATA_VALIDATION_PLUGIN");
        entry.put("data", "{\"s1\":{\"rule1\":{\"uid\":\"rule1\",\"type\":\"list\",\"formula1\":\"Yes,No\",\"showDropDown\":true,\"ranges\":[{\"startRow\":0,\"startColumn\":0,\"endRow\":1,\"endColumn\":0,\"startAbsoluteRefType\":0,\"endAbsoluteRefType\":0,\"rangeType\":0}]}}}");
        resources.add(entry);
        src.setResources(resources);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        UniverXlsx.write(src, out, UniverXlsxOptions.builder().writeSidecar(false).build());

        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()), UniverXlsxOptions.builder().writeSidecar(false).build());
        assertThat(back.getResources()).isInstanceOf(List.class);
        String data = null;
        for (Object item : (List<?>) back.getResources()) {
            if (item instanceof Map && "SHEET_DATA_VALIDATION_PLUGIN".equals(String.valueOf(((Map<?, ?>) item).get("name")))) {
                data = String.valueOf(((Map<?, ?>) item).get("data"));
                break;
            }
        }
        assertThat(data).contains("Yes,No");
        assertThat(data).contains("\"type\":\"list\"");
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
