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
package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsxException;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.UniverXlsxReadException;
import io.github.autoffice.univer.converter.WorkbookConverter;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.resource.SidecarPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;

/**
 * xlsx → IWorkbookData 读取器。
 * xlsx reader producing IWorkbookData.
 */
public class UniverXlsxReader {
    private final UniverXlsxOptions opts;

    public UniverXlsxReader(UniverXlsxOptions opts) {
        this.opts = opts;
    }

    /**
     * 读取 xlsx 为 IWorkbookData。
     * Read xlsx as IWorkbookData.
     */
    public IWorkbookData read(InputStream in) throws UniverXlsxException {
        try (OPCPackage pkg = OPCPackage.open(in)) {
            Optional<IWorkbookData> sidecar = SidecarPart.read(pkg);
            try (XSSFWorkbook wb = new XSSFWorkbook(pkg)) {
                return new WorkbookConverter(opts).fromXlsx(wb, sidecar.orElse(null));
            }
        } catch (IOException e) {
            throw new UniverXlsxReadException("read xlsx failed: " + e.getMessage(), e);
        } catch (Exception e) {
            throw new UniverXlsxReadException("read xlsx failed: " + e.getMessage(), e);
        }
    }
}
