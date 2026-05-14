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
package io.github.autoffice.univer.resource;

import io.github.autoffice.univer.UniverXlsxWriteException;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Collections;
import java.util.Optional;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SidecarPartTest {

    @Test
    void should_roundtrip_sidecar_part() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s1");
            IWorkbookData meta = new IWorkbookData().setId("w").setAppVersion("0.10.2").setLocale("enUS");
            meta.getSheets().put("s1", new IWorksheetData().setId("s1").setName("s1"));
            meta.setSheetOrder(Collections.singletonList("s1"));
            SidecarPart.write(wb.getPackage(), meta, false);
            wb.write(buf);
        }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            Optional<IWorkbookData> read = SidecarPart.read(pkg);
            assertThat(read).isPresent();
            assertThat(read.get().getAppVersion()).isEqualTo("0.10.2");
            assertThat(read.get().getId()).isEqualTo("w");
            assertThat(read.get().getSheets()).containsKey("s1");
        }
    }

    @Test
    void should_return_empty_when_sidecar_missing() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s1");
            wb.write(buf);
        }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            assertThat(SidecarPart.read(pkg)).isEmpty();
        }
    }

    @Test
    void should_write_pretty_json_when_requested() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s1");
            IWorkbookData meta = new IWorkbookData().setId("w").setAppVersion("0.10.2");
            SidecarPart.write(wb.getPackage(), meta, true);
            wb.write(buf);
        }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            Optional<IWorkbookData> read = SidecarPart.read(pkg);
            assertThat(read).isPresent();
            assertThat(read.get().getId()).isEqualTo("w");
        }
    }

    @Test
    void should_overwrite_existing_sidecar() throws Exception {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s1");
            SidecarPart.write(wb.getPackage(), new IWorkbookData().setId("v1"), false);
            SidecarPart.write(wb.getPackage(), new IWorkbookData().setId("v2"), false);
            wb.write(buf);
        }
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(buf.toByteArray()))) {
            Optional<IWorkbookData> read = SidecarPart.read(pkg);
            assertThat(read).isPresent();
            assertThat(read.get().getId()).isEqualTo("v2");
        }
    }

    @Test
    void should_throw_UniverXlsxWriteException_on_write_failure() {
        // 传入 null 包以触发 NPE，验证异常被正确包装为 UniverXlsxWriteException。
        // Passing a null package triggers NPE, verifying wrapping into UniverXlsxWriteException.
        assertThatThrownBy(() -> SidecarPart.write(null, new IWorkbookData().setId("x"), false))
                .isInstanceOf(UniverXlsxWriteException.class)
                .hasMessageContaining("sidecar");
    }
}
