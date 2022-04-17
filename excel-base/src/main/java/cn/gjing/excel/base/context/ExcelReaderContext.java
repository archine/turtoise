package cn.gjing.excel.base.context;

import cn.gjing.excel.base.listener.read.ExcelResultReadListener;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Excel reader global context
 *
 * @author Gjing
 **/
@Getter
@Setter
public class ExcelReaderContext<R> extends AbstractExcelContext {
    /**
     * Header names for an Excel file
     */
    private List<String> headNames;

    /**
     * Current excel entity
     */
    private Class<R> excelEntity;

    /**
     * Excel header mapping field
     */
    private Map<String, Field> excelFieldMap;

    /**
     * Check that the Excel file is bound to the currently set mapping entity
     */
    private boolean checkTemplate = false;

    /**
     * Read rows before the header
     */
    private boolean readOther = false;

    /**
     * The unique key
     */
    private String uniqueKey;

    /**
     * Ignore the array of actual Excel table headers that you read when importing
     */
    private String[] ignores;

    /**
     * Read result listener
     */
    private ExcelResultReadListener<R> resultReadListener;

    public ExcelReaderContext() {
        super();
    }

    public ExcelReaderContext(Class<R> excelEntity, Map<String, Field> excelFieldMap, String[] ignores) {
        super();
        this.excelEntity = excelEntity;
        this.excelFieldMap = excelFieldMap;
        this.headNames = new ArrayList<>();
        this.ignores = ignores;
        this.checkTemplate = false;
        this.readOther = false;
    }
}
