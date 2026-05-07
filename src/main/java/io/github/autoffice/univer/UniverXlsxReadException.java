package io.github.autoffice.univer;

/**
 * 读取 xlsx 时抛出的异常。
 * Thrown when reading xlsx fails.
 */
public class UniverXlsxReadException extends UniverXlsxException {
    public UniverXlsxReadException(String msg) {
        super(msg);
    }

    public UniverXlsxReadException(String msg, Throwable cause) {
        super(msg, cause);
    }
}
