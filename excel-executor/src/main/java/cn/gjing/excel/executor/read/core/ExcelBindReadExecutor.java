package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.annotation.ExcelAssert;
import cn.gjing.excel.base.annotation.ExcelDataConvert;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelAssertException;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.ExcelListener;
import cn.gjing.excel.base.meta.ELMeta;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ListenerChain;
import cn.gjing.excel.base.util.ParamUtils;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.executor.util.JsonUtils;
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
 * Excel bind mode import executor
 * @author Gjing
 **/
public class ExcelBindReadExecutor<R> extends ExcelBaseReadExecutor<R> {

    public ExcelBindReadExecutor(ExcelReaderContext<R> context) {
        super(context);
    }

    @Override
    public void read(int headerIndex, String sheetName) {
        super.validTemplate();
        super.checkSheet(sheetName);
        this.reader(headerIndex, super.context.getResultReadListener() == null ? null : new ArrayList<>(), super.context.getListenerCache(), new StandardEvaluationContext());
    }

    /**
     * Start read
     *
     * @param headerIndex Excel header index
     * @param dataList    All data
     */
    private void reader(int headerIndex, List<R> dataList, List<ExcelListener> rowReadListeners, EvaluationContext context) {
        R r;
        boolean continueRead = true;
        ListenerChain.doReadBefore(rowReadListeners);
        for (Row row : super.context.getSheet()) {
            if (!continueRead) {
                break;
            }
            if (row.getRowNum() < headerIndex) {
                continueRead = super.readOther(rowReadListeners, row);
                continue;
            }
            if (row.getRowNum() == headerIndex) {
                continueRead = super.readHeader(rowReadListeners, row);
                continue;
            }
            super.saveCurrentRowObj = true;
            try {
                r = this.context.getExcelEntity().newInstance();
                context.setVariable(super.context.getExcelEntity().getSimpleName(), r);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new ExcelException("Excel entity init failure, " + e.getMessage());
            }
            for (int c = 0, size = super.context.getHeadNames().size(); c < size; c++) {
                String head = super.context.getHeadNames().get(c);
                if ("ignored".equals(head)) {
                    continue;
                }
                Field field = super.context.getExcelFieldMap().get(head);
                if (field == null) {
                    field = super.context.getExcelFieldMap().get(head + ParamUtils.numberToEn(c));
                }
                if (field == null) {
                    continue;
                }
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                Cell valueCell = row.getCell(c + super.startCol);
                Object value;
                if (valueCell != null) {
                    value = super.getValue(r, valueCell, field, excelField.trim(), excelField.required(), RowType.BODY);
                    if (!super.saveCurrentRowObj) {
                        break;
                    }
                    context.setVariable(field.getName(), value);
                    this.assertValue(context, row, c, field, excelField);
                    value = this.convert(value , context, field.getAnnotation(ExcelDataConvert.class));
                    value = ListenerChain.doReadCell(rowReadListeners, value, valueCell, row.getRowNum(), c, RowType.BODY);
                } else {
                    if (excelField.required()) {
                        super.saveCurrentRowObj = ListenerChain.doReadEmpty(this.context.getListenerCache(), r, row.getRowNum(), c);
                        if (!super.saveCurrentRowObj) {
                            break;
                        }
                    }
                    context.setVariable(field.getName(), null);
                    this.assertValue(context, row, c, field, excelField);
                    value = this.convert(null, context, field.getAnnotation(ExcelDataConvert.class));
                    value = ListenerChain.doReadCell(rowReadListeners, value, null, row.getRowNum(), c, RowType.BODY);
                }
                if (value != null) {
                    this.setValue(r, field, value);
                }
                context.setVariable(field.getName(), value);
            }
            if (super.saveCurrentRowObj) {
                continueRead = ListenerChain.doReadRow(rowReadListeners, r, row, RowType.BODY);
                if (dataList != null) {
                    dataList.add(r);
                }
            }
        }
        ListenerChain.doReadFinish(rowReadListeners);
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
            return ELMeta.PARSER.getParser().parseExpression(excelDataConvert.readExpr()).getValue(context);
        }
        return value;
    }

    /**
     * Set value for the field of the object
     *
     * @param o     object
     * @param field field
     * @param value value
     */
    private void setValue(R o, Field field, Object value) {
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
            throw new ExcelException("Unsupported data type, the current cell value type is " + value.getClass().getTypeName() + ", but " + field.getName() + " is " + field.getType().getTypeName());
        }
    }

    /**
     * Cell value assert
     *
     * @param context    EL context
     * @param row        Current row
     * @param c          Current col index
     * @param field      Current field
     * @param excelField ExcelFiled annotation on current filed
     */
    private void assertValue(EvaluationContext context, Row row, int c, Field field, ExcelField excelField) {
        ExcelAssert excelAssert = field.getAnnotation(ExcelAssert.class);
        if (excelAssert != null) {
            Boolean test = ELMeta.PARSER.getParser().parseExpression(excelAssert.expr()).getValue(context, Boolean.class);
            if (test != null && !test) {
                throw new ExcelAssertException(excelAssert.message(), excelField, field, row.getRowNum(), c);
            }
        }
    }
}
