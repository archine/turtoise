package cn.gjing.excel.base.context;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.meta.ExcelType;
import lombok.Getter;
import lombok.Setter;

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
    private short headerHeight = 450;

    /**
     * Excel body row height
     */
    private short bodyHeight = 390;

    /**
     * Excel header series
     */
    private int headerSeries = 1;

    /**
     * Whether the current export is a template
     */
    private boolean isTemplate;

    /**
     * Excel filed properties
     */
    private List<ExcelFieldProperty> fieldProperties;

    public ExcelWriterContext() {
        super();
    }
}
