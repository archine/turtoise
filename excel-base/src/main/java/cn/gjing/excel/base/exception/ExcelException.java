package cn.gjing.excel.base.exception;

/**
 * Excel exception parent
 *
 * @author Gjing
 **/
public class ExcelException extends RuntimeException {
    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
    }
}
