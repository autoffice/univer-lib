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

import io.github.autoffice.univer.io.UniverXlsxReader;
import io.github.autoffice.univer.io.UniverXlsxWriter;
import io.github.autoffice.univer.model.IWorkbookData;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * xlsx ↔ IWorkbookData 转换门面。
 * Facade for xlsx ↔ IWorkbookData conversion.
 */
public final class UniverXlsx {
    private UniverXlsx() {
    }

    /**
     * 从输入流读取为 IWorkbookData。
     * Read xlsx from input stream into IWorkbookData.
     */
    public static IWorkbookData read(InputStream in) throws UniverXlsxException {
        return read(in, UniverXlsxOptions.defaults());
    }

    /**
     * 从路径读取为 IWorkbookData。
     * Read xlsx from path into IWorkbookData.
     */
    public static IWorkbookData read(Path path) throws UniverXlsxException {
        try (InputStream in = Files.newInputStream(path)) {
            return read(in, UniverXlsxOptions.defaults());
        } catch (IOException e) {
            throw new UniverXlsxReadException("open file failed: " + path, e);
        }
    }

    /**
     * 使用自定义选项读取。
     * Read with custom options.
     */
    public static IWorkbookData read(InputStream in, UniverXlsxOptions opts) throws UniverXlsxException {
        return new UniverXlsxReader(opts).read(in);
    }

    /**
     * 将 IWorkbookData 写入输出流。
     * Write IWorkbookData to output stream as xlsx.
     */
    public static void write(IWorkbookData wb, OutputStream out) throws UniverXlsxException {
        write(wb, out, UniverXlsxOptions.defaults());
    }

    /**
     * 将 IWorkbookData 写入文件。
     * Write IWorkbookData to path as xlsx.
     */
    public static void write(IWorkbookData wb, Path path) throws UniverXlsxException {
        try (OutputStream out = Files.newOutputStream(path)) {
            write(wb, out, UniverXlsxOptions.defaults());
        } catch (IOException e) {
            throw new UniverXlsxWriteException("open file failed: " + path, e);
        }
    }

    /**
     * 使用自定义选项写入。
     * Write with custom options.
     */
    public static void write(IWorkbookData wb, OutputStream out, UniverXlsxOptions opts) throws UniverXlsxException {
        new UniverXlsxWriter(opts).write(wb, out);
    }
}
