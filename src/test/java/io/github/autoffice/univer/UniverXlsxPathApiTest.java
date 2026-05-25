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
package io.github.autoffice.univer;

import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * 覆盖 UniverXlsx 基于 Path 的 read/write 入口和 InputStream 重载。
 * Covers the Path-based read/write API methods that were not previously exercised.
 */
class UniverXlsxPathApiTest {

    @Test
    void should_write_then_read_via_path(@TempDir Path tmp) throws Exception {
        IWorkbookData src = new IWorkbookData().setId("pid").setAppVersion("0.1.0");
        src.getSheets().put("s1", new IWorksheetData().setId("s1").setName("Sheet1"));
        src.setSheetOrder(Collections.singletonList("s1"));
        Path file = tmp.resolve("out.xlsx");
        UniverXlsx.write(src, file);
        assertThat(Files.exists(file)).isTrue();
        assertThat(Files.size(file)).isPositive();

        IWorkbookData back = UniverXlsx.read(file);
        assertThat(back.getId()).isEqualTo("pid");
        assertThat(back.getSheets()).containsKey("s1");
    }

    @Test
    void should_read_inputstream_with_default_options() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("isr");
        src.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S"));
        src.setSheetOrder(Collections.singletonList("s1"));
        java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
        UniverXlsx.write(src, out);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()));
        assertThat(back.getId()).isEqualTo("isr");
    }

    @Test
    void should_throw_read_exception_for_missing_path(@TempDir Path tmp) {
        Path missing = tmp.resolve("does-not-exist.xlsx");
        assertThatThrownBy(() -> UniverXlsx.read(missing))
                .isInstanceOf(UniverXlsxReadException.class)
                .hasMessageContaining("open file failed");
    }

    @Test
    void should_throw_write_exception_for_unwritable_path(@TempDir Path tmp) {
        // 不存在的子目录路径 / parent directory does not exist
        Path bad = tmp.resolve("nope/does-not-exist/out.xlsx");
        IWorkbookData src = new IWorkbookData();
        src.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S"));
        src.setSheetOrder(Collections.singletonList("s1"));
        assertThatThrownBy(() -> UniverXlsx.write(src, bad))
                .isInstanceOf(UniverXlsxWriteException.class)
                .hasMessageContaining("open file failed");
    }

    @Test
    void should_skip_sidecar_when_disabled() throws Exception {
        IWorkbookData src = new IWorkbookData().setId("nosc").setAppVersion("0.2.0");
        src.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S"));
        src.setSheetOrder(Collections.singletonList("s1"));
        UniverXlsxOptions opts = UniverXlsxOptions.builder().writeSidecar(false).build();
        java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
        UniverXlsx.write(src, out, opts);
        IWorkbookData back = UniverXlsx.read(new ByteArrayInputStream(out.toByteArray()), opts);
        // 没有边车则 appVersion 丢失，并且 sheet 以 sheetName 作为 id 兜底
        // No sidecar means sheet id falls back to sheet name, and appVersion (sidecar-only) is gone
        assertThat(back.getAppVersion()).isNull();
        assertThat(back.getSheets()).containsKey("S");
    }
}
