package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsxException;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.UniverXlsxWriteException;
import io.github.autoffice.univer.model.IWorkbookData;

import java.io.OutputStream;

/**
 * xlsx 写出器（空壳实现，Task 13 装配）。
 * xlsx writer (stub, wired in Task 13).
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
        throw new UniverXlsxWriteException("not implemented");
    }
}
