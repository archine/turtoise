package cn.gjing.excel.valid.listener;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelCellWriteListener;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.valid.handler.HandleMeta;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Map;

/**
 * Validates annotation processing listener,
 * used for logical processing of validation annotations
 *
 * @author Gjing
 **/
public class ValidAnnotationProcessListener implements ExcelCellWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext context;
    private final Map<String, String[]> boxValues;
    private final Map<String, String[]> cascadeValues;

    public ValidAnnotationProcessListener() {
        this.boxValues = null;
        this.cascadeValues = null;
    }

    /**
     * You can set the options of the drop-down box by using the method,
     * which is used when the drop-down box has too many options
     *
     * @param boxValues drop-down box or cascade box options
     */
    public ValidAnnotationProcessListener(Map<String, String[]> boxValues) {
        this.boxValues = boxValues;
        this.cascadeValues = null;
    }

    /**
     * You can set the options of the drop-down box by using the method,
     * which is used when the drop-down box has too many options or cascading
     *
     * @param boxValues drop-down box or cascade box options
     * @param cascadeValues Cascade drop-down box values
     */
    public ValidAnnotationProcessListener(Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        this.boxValues = boxValues;
        this.cascadeValues = cascadeValues;
    }

    @Override
    public void setContext(ExcelWriterContext writerContext) {
        this.context = writerContext;
    }

    @Override
    public void completeCell(Sheet sheet, Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex, RowType rowType) {
        if (rowType == RowType.HEAD) {
            if (dataIndex + 1 == this.context.getHeaderSeries()) {
                HandleMeta.INSTANCE.exec(property.getField(), this.context, row, colIndex, this.boxValues,this.cascadeValues);
            }
        }
    }
}
