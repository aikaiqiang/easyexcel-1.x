package com.alibaba.excel.exception;

/**
 *
 * @author jipengfei
 */
public class ExcelExitException extends RuntimeException {

    public ExcelExitException() {
    }

    public ExcelExitException(String message) {
        super(message);
    }

    public ExcelExitException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelExitException(Throwable cause) {
        super(cause);
    }
}
