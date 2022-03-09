package cn.gjing.excel.executor.write.context;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.context.AbstractExcelContext;
import cn.gjing.excel.base.meta.ExcelType;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel writer global context
 *
 * @author Gjing
 **/
@Getter
@Setter
public final class ExcelWriterContext extends AbstractExcelContext {
    /**
     * Current excel mapping entity
     */
    private Class<?> excelEntity;

    /**
     * Excel file name
     */
    private String fileName;

    /**
     * Whether to open multistage Excel headers
     */
    private boolean multiHead = false;

    /**
     * Whether a header exists
     */
    private boolean existHead = true;

    /**
     * Whether is excel template file
     */
    private boolean isTemplate = false;

    /**
     * Excel header fields
     */
    private List<Field> excelFields = new ArrayList<>();

    /**
     * Whether you need to add a file identifier when exporting an Excel file,
     * which can be used for template validation of the file at import time
     */
    private boolean bind = true;

    /**
     * The unique key
     */
    private String uniqueKey;

    /**
     * Excel type
     */
    private ExcelType excelType = ExcelType.XLS;

    /**
     * Excel head row height
     */
    private short headerHeight = 400;

    /**
     * Excel body row height
     */
    private short bodyHeight = 370;

    /**
     * Excel header series
     */
    private int headerSeries = 1;

    /**
     * Excel filed property
     */
    private List<ExcelFieldProperty> fieldProperties;

    public ExcelWriterContext() {
        super();
    }
}
