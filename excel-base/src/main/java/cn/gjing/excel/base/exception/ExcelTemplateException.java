package cn.gjing.excel.base.exception;

/**
 * Excel template matching exception.
 * The exception will be thrown if the template fails to match the unique ID of the Excel entity during import.
 * The exception will be thrown even if the template is not an Excel file during import
 *
 * @author Gjing
 **/
public class ExcelTemplateException extends ExcelException{
    public ExcelTemplateException() {
        super("Excel template do not match");
    }

    public ExcelTemplateException(String message) {
        super(message);
    }
}
