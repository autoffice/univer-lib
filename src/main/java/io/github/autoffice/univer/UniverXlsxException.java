package io.github.autoffice.univer;

import java.io.IOException;

/**
 * 库顶层受检异常。
 * Top-level checked exception for all library operations.
 */
public class UniverXlsxException extends IOException {
    public UniverXlsxException(String msg) {
        super(msg);
    }

    public UniverXlsxException(String msg, Throwable cause) {
        super(msg, cause);
    }
}
