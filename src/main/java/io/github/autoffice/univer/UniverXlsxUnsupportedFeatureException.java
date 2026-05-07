package io.github.autoffice.univer;

/**
 * 严格模式下遇到不支持特性时抛出。
 * Thrown in strict mode when an unsupported feature is encountered.
 */
public class UniverXlsxUnsupportedFeatureException extends UniverXlsxException {
    public UniverXlsxUnsupportedFeatureException(String msg) {
        super(msg);
    }
}
