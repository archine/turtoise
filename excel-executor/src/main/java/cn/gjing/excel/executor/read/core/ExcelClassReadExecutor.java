package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelAssert;
import cn.gjing.excel.base.annotation.ExcelDataConvert;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelAssertException;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.meta.ELMeta;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.meta.WRMode;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.executor.util.JsonUtils;
import cn.gjing.excel.executor.util.ListenerChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Import data to generate the specified class object.
 * Each row of data corresponds to an object
 *
 * @author Gjing
 **/
public class ExcelClassReadExecutor<R> extends ExcelBaseReadExecutor<R> {

    public ExcelClassReadExecutor(ExcelReaderContext<R> context) {
        super(context);
    }

    @Override
    public void read(int headerIndex, String sheetName) {
        super.validTemplate();
        super.checkSheet(sheetName);
        this.reader(headerIndex, super.context.getResultReadListener() == null ? null : new ArrayList<>(), new StandardEvaluationContext());
    }

    /**
     * Start read
     *
     * @param headerIndex Excel header index
     * @param dataList    All data
     */
    private void reader(int headerIndex, List<R> dataList, EvaluationContext context) {
        R r;
        boolean continueRead = true;
        ListenerChain.doReadBefore(super.context.getListenerCache());
        for (Row row : super.context.getSheet()) {
            if (!continueRead) {
                break;
            }
            int rowNum = row.getRowNum();
            if (rowNum < headerIndex) {
                continueRead = super.readOther(row);
                continue;
            }
            if (rowNum == headerIndex) {
                continueRead = super.readHeader(row);
                continue;
            }
            super.saveCurrentRowObj = true;
            try {
                r = this.context.getExcelEntity().newInstance();
                context.setVariable(super.context.getExcelEntity().getSimpleName(), r);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new ExcelException("Excel entity init failure, " + e.getMessage());
            }
            for (int fieldIndex = 0, size = super.context.getFieldProperties().size(); fieldIndex < size; fieldIndex++) {
                ExcelFieldProperty property = this.context.getFieldProperties().get(fieldIndex);
                int colNum = super.context.getWrMode() == WRMode.INDEX ? property.getIndex() : fieldIndex;
                ExcelField excelField = property.getField().getAnnotation(ExcelField.class);
                Cell valueCell = row.getCell(colNum);
                Object value;
                if (valueCell != null) {
                    value = super.getValue(r, valueCell, excelField.trim(), excelField.required());
                    if (!super.saveCurrentRowObj) {
                        break;
                    }
                    context.setVariable(property.getField().getName(), value);
                    this.assertValue(context, row, colNum, property.getField(), excelField);
                    value = this.convert(value, context, property.getField().getAnnotation(ExcelDataConvert.class));
                    value = ListenerChain.doReadCell(super.context.getListenerCache(), value, valueCell, rowNum, colNum, RowType.BODY);
                } else {
                    if (excelField.required()) {
                        super.saveCurrentRowObj = ListenerChain.doReadEmpty(this.context.getListenerCache(), r, rowNum, colNum);
                        if (!super.saveCurrentRowObj) {
                            break;
                        }
                    }
                    context.setVariable(property.getField().getName(), null);
                    this.assertValue(context, row, colNum, property.getField(), excelField);
                    value = this.convert(null, context, property.getField().getAnnotation(ExcelDataConvert.class));
                    value = ListenerChain.doReadCell(super.context.getListenerCache(), value, null, rowNum, colNum, RowType.BODY);
                }
                if (value != null) {
                    this.setValue(r, property.getField(), value, rowNum, colNum);
                }
                context.setVariable(property.getField().getName(), value);
            }
            if (super.saveCurrentRowObj) {
                continueRead = ListenerChain.doReadRow(super.context.getListenerCache(), r, row, RowType.BODY);
                if (dataList != null) {
                    dataList.add(r);
                }
            }
        }
        ListenerChain.doReadFinish(super.context.getListenerCache());
        if (this.context.getResultReadListener() != null) {
            this.context.getResultReadListener().notify(dataList);
        }
    }

    /**
     * Data convert
     *
     * @param value            Attribute values
     * @param excelDataConvert excelDataConvert
     * @param context          EL context
     * @return new value
     */
    private Object convert(Object value, EvaluationContext context, ExcelDataConvert excelDataConvert) {
        if (excelDataConvert != null && !"".equals(excelDataConvert.readExpr())) {
            return ELMeta.PARSER.parse(excelDataConvert.readExpr(), context);
        }
        return value;
    }

    /**
     * Set value for the field of the object
     *
     * @param o        object
     * @param field    field
     * @param value    value
     * @param colIndex current col index
     * @param rowIndex current row index
     */
    private void setValue(R o, Field field, Object value, int rowIndex, int colIndex) {
        try {
            if (field.getType() != value.getClass()) {
                value = JsonUtils.toObj(JsonUtils.toJson(value), field.getType());
            }
            BeanUtils.setFieldValue(o, field, value);
        } catch (RuntimeException e) {
            if (field.getType() == LocalDate.class) {
                BeanUtils.setFieldValue(o, field, LocalDateTime.ofInstant(((Date) value).toInstant(), ZoneId.systemDefault()).toLocalDate());
                return;
            }
            if (field.getType() == LocalDateTime.class) {
                BeanUtils.setFieldValue(o, field, LocalDateTime.ofInstant(((Date) value).toInstant(), ZoneId.systemDefault()));
                return;
            }
            throw new ExcelException("unsupported data type, the current cell" + "[row:" + rowIndex + ",column:" + colIndex + "]" + " value type is " + value.getClass().getTypeName() + ", but " + field.getName() + " is " + field.getType().getTypeName());
        }
    }

    /**
     * Cell value assert
     *
     * @param context    EL context
     * @param row        Current row
     * @param colIndex   Current col index
     * @param field      Current field
     * @param excelField ExcelFiled annotation on current filed
     */
    private void assertValue(EvaluationContext context, Row row, int colIndex, Field field, ExcelField excelField) {
        ExcelAssert excelAssert = field.getAnnotation(ExcelAssert.class);
        if (excelAssert != null) {
            Boolean test = ELMeta.PARSER.parse(excelAssert.expr(), context, Boolean.class);
            if (test != null && !test) {
                throw new ExcelAssertException(excelAssert.message(), excelField, field, row.getRowNum(), colIndex);
            }
        }
    }
}
