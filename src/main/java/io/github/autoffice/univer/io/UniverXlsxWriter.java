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
import io.github.autoffice.univer.UniverXlsxWriteException;
import io.github.autoffice.univer.converter.WorkbookConverter;
import io.github.autoffice.univer.model.IWorkbookData;
import io.github.autoffice.univer.resource.SidecarPart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

/**
 * IWorkbookData → xlsx 写出器。
 * xlsx writer consuming IWorkbookData.
 */
public class UniverXlsxWriter {
    private final UniverXlsxOptions opts;

    public UniverXlsxWriter(UniverXlsxOptions opts) {
        this.opts = opts;
    }

    /**
     * 写出 IWorkbookData 为 xlsx。
     * Write IWorkbookData as xlsx.
     */
    public void write(IWorkbookData wb, OutputStream out) throws UniverXlsxException {
        try (XSSFWorkbook poiWb = new WorkbookConverter(opts).toXlsx(wb)) {
            if (opts.isWriteSidecar() && wb != null) {
                SidecarPart.write(poiWb.getPackage(), wb, opts.isPrettyJson());
            }
            poiWb.write(out);
        } catch (IOException e) {
            throw new UniverXlsxWriteException("write xlsx failed: " + e.getMessage(), e);
        } catch (RuntimeException e) {
            // POI 在保存 OPC 包时常把 IOException 包成 OpenXML4JRuntimeException 等非受检异常，
            // 同样需要包装为库的受检异常以保持公共 API 契约。
            throw new UniverXlsxWriteException("write xlsx failed: " + e.getMessage(), e);
        }
    }
}
