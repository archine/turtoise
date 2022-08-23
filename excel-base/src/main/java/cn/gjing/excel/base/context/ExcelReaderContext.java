package cn.gjing.excel.base.context;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.listener.read.ExcelResultReadListener;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

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
    private List<Object> headNames;

    /**
     * Current excel entity
     */
    private Class<R> excelEntity;

    /**
     * Check that the Excel file is bound to the currently set mapping entity
     */
    private boolean checkTemplate = false;

    /**
     * Read rows before the header
     */
    private boolean readOther = false;

    /**
     * The Excel id card
     */
    private String idCard;

    /**
     * Read result listener
     */
    private ExcelResultReadListener<R> resultReadListener;

    /**
     * Excel filed properties
     */
    private List<ExcelFieldProperty> fieldProperties;

    public ExcelReaderContext() {
        super();
    }

    public ExcelReaderContext(Class<R> excelEntity) {
        super();
        this.excelEntity = excelEntity;
        this.headNames = new ArrayList<>();
        this.checkTemplate = false;
        this.readOther = false;
    }
}
