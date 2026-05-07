package io.github.autoffice.univer.io;

import io.github.autoffice.univer.UniverXlsxException;
import io.github.autoffice.univer.UniverXlsxOptions;
import io.github.autoffice.univer.UniverXlsxReadException;
import io.github.autoffice.univer.model.IWorkbookData;

import java.io.InputStream;

/**
 * xlsx 读取器（空壳实现，Task 13 装配）。
 * xlsx reader (stub, wired in Task 13).
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
        throw new UniverXlsxReadException("not implemented");
    }
}
