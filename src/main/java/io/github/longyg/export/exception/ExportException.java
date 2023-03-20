package io.github.longyg.export.exception;

/**
 * @author longyg
 */
public class ExportException extends Exception {
    public ExportException(Throwable e) {
        super(e);
    }

    public ExportException(String msg, Throwable e) {
        super(msg, e);
    }
}
