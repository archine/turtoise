package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelDataConvert;
import cn.gjing.excel.base.meta.ELMeta;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.executor.util.ExcelUtils;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * Export processor for Excel bind mode
 *
 * @author Gjing
 **/
class ExcelBindWriterExecutor extends ExcelBaseWriteExecutor {
    public ExcelBindWriterExecutor(ExcelWriterContext context) {
        super(context);
    }

    /**
     * Set excel head
     *
     * @param needHead  Whether to set header
     * @param boxValues Excel dropdown box value
     */
    @Override
    public void writeHead(boolean needHead, Map<String, String[]> boxValues) {
        if (!needHead || this.context.getFieldProperties().isEmpty()) {
            this.context.setExistHead(false);
            return;
        }
        Row headRow;
        String[] currentRowHeadArray = new String[this.context.getFieldProperties().size()];
        for (int index = 0; index < this.context.getHeaderSeries(); index++) {
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), index, RowType.HEAD);
            headRow = this.context.getSheet().createRow(this.context.getSheet().getPhysicalNumberOfRows());
            if (this.context.getHeaderHeight() > 0) {
                headRow.setHeight(this.context.getHeaderHeight());
            }
            for (int colIndex = 0, headSize = this.context.getFieldProperties().size(); colIndex < headSize; colIndex++) {
                ExcelFieldProperty property = this.context.getFieldProperties().get(colIndex);
                String headName = property.getValue()[index];
                currentRowHeadArray[colIndex] = headName;
                Cell headCell = headRow.createCell(headRow.getPhysicalNumberOfCells());
                ListenerChain.doSetHeadStyle(this.context.getListenerCache(), headRow, headCell, property, index, colIndex);
                headName = (String) ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), headRow, headCell, property, index, headCell.getColumnIndex(), RowType.HEAD, headName);
                headCell.setCellValue(headName);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), headRow, headCell, property, index, headCell.getColumnIndex(), RowType.HEAD);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), headRow, currentRowHeadArray, index, RowType.HEAD);
        }
    }

    /**
     * Set excel body
     *
     * @param data Export data
     */
    @Override
    public void writeBody(List<?> data) {
        EvaluationContext context = new StandardEvaluationContext();
        for (int index = 0, dataSize = data.size(); index < dataSize; index++) {
            Object o = data.get(index);
            context.setVariable(o.getClass().getSimpleName(), o);
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), index, RowType.BODY);
            Row valueRow = this.context.getSheet().createRow(this.context.getSheet().getPhysicalNumberOfRows());
            if (this.context.getBodyHeight() > 0) {
                valueRow.setHeight(this.context.getBodyHeight());
            }
            for (int colIndex = 0, headSize = this.context.getExcelFields().size(); colIndex < headSize; colIndex++) {
                Field field = this.context.getExcelFields().get(colIndex);
                ExcelFieldProperty property = this.context.getFieldProperties().get(colIndex);
                Object value = BeanUtils.getFieldValue(o, field);
                Cell valueCell = valueRow.createCell(valueRow.getPhysicalNumberOfCells());
                context.setVariable(field.getName(), value);
                ListenerChain.doSetBodyStyle(this.context.getListenerCache(), valueRow, valueCell, property, index, colIndex);
                value = this.convert(value, field.getAnnotation(ExcelDataConvert.class), context);
                value = ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property, index, valueCell.getColumnIndex(), RowType.BODY, value);
                ExcelUtils.setCellValue(valueCell, value);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property, index, valueCell.getColumnIndex(), RowType.BODY);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), valueRow, o, index, RowType.BODY);
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
    private Object convert(Object value, ExcelDataConvert excelDataConvert, EvaluationContext context) {
        if (excelDataConvert != null && !"".equals(excelDataConvert.writeExpr())) {
            return ELMeta.PARSER.getParser().parseExpression(excelDataConvert.writeExpr()).getValue(context);
        }
        return value;
    }
}
