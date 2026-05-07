package io.github.autoffice.univer;

/**
 * 写出 xlsx 时抛出的异常。
 * Thrown when writing xlsx fails.
 */
public class UniverXlsxWriteException extends UniverXlsxException {
    public UniverXlsxWriteException(String msg) {
        super(msg);
    }

    public UniverXlsxWriteException(String msg, Throwable cause) {
        super(msg, cause);
    }
}
