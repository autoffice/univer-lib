/*
 * Copyright © 2026 AutOffice (hello.aldis@qq.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.UniverXlsxReadException;
import io.github.autoffice.univer.UniverXlsxWriteException;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.model.IWorksheetData;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collections;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * 覆盖 UniverXlsxReader / UniverXlsxWriter 的 IOException 包装路径。
 */
class UniverXlsxIoErrorTest {

    @Test
    void should_wrap_io_exception_in_writer() {
        IWorkbookData wb = new IWorkbookData();
        wb.getSheets().put("s1", new IWorksheetData().setId("s1").setName("S"));
        wb.setSheetOrder(Collections.singletonList("s1"));
        OutputStream bad = new OutputStream() {
            @Override
            public void write(int b) throws IOException {
                throw new IOException("boom");
            }
            @Override
            public void write(byte[] b, int off, int len) throws IOException {
                throw new IOException("boom");
            }
        };
        UniverXlsxWriter writer = new UniverXlsxWriter(UniverXlsxOptions.defaults());
        assertThatThrownBy(() -> writer.write(wb, bad))
                .isInstanceOf(UniverXlsxWriteException.class)
                .hasMessageContaining("write xlsx failed");
    }

    @Test
    void should_wrap_io_exception_in_reader() {
        // 让输入流读取到一半就抛 IOException
        InputStream bad = new InputStream() {
            int n = 0;
            @Override
            public int read() throws IOException {
                if (n++ < 10) {
                    return 0;
                }
                throw new IOException("boom");
            }
        };
        UniverXlsxReader reader = new UniverXlsxReader(UniverXlsxOptions.defaults());
        assertThatThrownBy(() -> reader.read(bad))
                .isInstanceOf(UniverXlsxReadException.class)
                .hasMessageContaining("read xlsx failed");
    }
}
