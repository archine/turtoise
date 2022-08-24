package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelDataConvert;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.meta.ELMeta;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.meta.WRMode;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.executor.util.ListenerChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.util.List;

/**
 * Export processor for Excel class mode
 *
 * @author Gjing
 **/
public class ExcelClassWriterExecutor extends ExcelBaseWriteExecutor {
    public ExcelClassWriterExecutor(ExcelWriterContext context) {
        super(context);
    }

    @Override
    public void writeBody(List<?> data) {
        EvaluationContext context = new StandardEvaluationContext();
        for (int dataIndex = 0, dataSize = data.size(); dataIndex < dataSize; dataIndex++) {
            Object o = data.get(dataIndex);
            context.setVariable(o.getClass().getSimpleName(), o);
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), dataIndex, RowType.BODY);
            Row valueRow = this.context.getSheet().createRow(this.context.getSheet().getLastRowNum() + 1);
            if (this.context.getBodyHeight() > 0) {
                valueRow.setHeight(this.context.getBodyHeight());
            }
            for (int fieldIndex = 0, headSize = this.context.getFieldProperties().size(); fieldIndex < headSize; fieldIndex++) {
                ExcelFieldProperty property = this.context.getFieldProperties().get(fieldIndex);
                Object value = BeanUtils.getFieldValue(o, property.getField());
                int lastCellNum = super.context.getWrMode() == WRMode.INDEX ? property.getIndex() : valueRow.getLastCellNum();
                Cell valueCell = valueRow.createCell(lastCellNum == -1 ? 0 : lastCellNum);
                context.setVariable(property.getField().getName(), value);
                ListenerChain.doSetBodyStyle(this.context.getListenerCache(), valueRow, valueCell, property, dataIndex);
                value = this.convert(value, property.getField().getAnnotation(ExcelDataConvert.class), context);
                value = ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property, dataIndex, RowType.BODY, value);
                ExcelUtils.setCellValue(valueCell, value);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property, dataIndex, RowType.BODY);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), valueRow, o, dataIndex, RowType.BODY);
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
            return ELMeta.PARSER.parse(excelDataConvert.writeExpr(), context);
        }
        return value;
    }
}
