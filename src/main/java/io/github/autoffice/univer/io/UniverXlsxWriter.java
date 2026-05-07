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
        }
    }
}
