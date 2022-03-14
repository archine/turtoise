package cn.gjing.excel.base.exception;

import cn.gjing.excel.base.annotation.ExcelAssert;
import cn.gjing.excel.base.annotation.ExcelField;
import lombok.Getter;

import java.lang.reflect.Field;

/**
 * Excel assert exception, thrown by {@link ExcelAssert}
 *
 * @author Gjing
 **/
@Getter
public class ExcelAssertException extends ExcelException {
    /**
     * ExcelFiled annotation on current filed
     */
    private final ExcelField excelField;
    /**
     * Current field
     */
    private final Field field;
    /**
     * Current row index from 0
     */
    private final int rowIndex;
    /**
     * Current column index from 0
     */
    private final int colIndex;

    public ExcelAssertException(String message, ExcelField excelField, Field field, int rowIndex, int colIndex) {
        super(message);
        this.excelField = excelField;
        this.field = field;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }
}
