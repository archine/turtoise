package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.executor.util.ExcelUtils;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.util.List;
import java.util.Map;

/**
 * Export processor for Excel simple mode
 *
 * @author Gjing
 **/
class ExcelSimpleWriterExecutor extends ExcelBaseWriteExecutor {
    public ExcelSimpleWriterExecutor(ExcelWriterContext context) {
        super(context);
    }

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
                String headName = this.context.getFieldProperties().get(colIndex).getValue()[index];
                currentRowHeadArray[colIndex] = headName;
                ExcelFieldProperty property = this.context.getFieldProperties().get(colIndex);
                Cell headCell = headRow.createCell(headRow.getPhysicalNumberOfCells());
                ListenerChain.doSetHeadStyle(this.context.getListenerCache(), headRow, headCell, property, index, colIndex);
                headName = (String) ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(),
                        headRow, headCell, property, index, headCell.getColumnIndex(), RowType.HEAD, headName);
                headCell.setCellValue(headName);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), headRow, headCell, property,
                        index, headCell.getColumnIndex(), RowType.HEAD);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), headRow, currentRowHeadArray,
                    index, RowType.HEAD);
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public void writeBody(List<?> data) {
        List<List<Object>> data2 = (List<List<Object>>) data;
        EvaluationContext context = new StandardEvaluationContext();
        for (int index = 0, dataSize = data.size(); index < dataSize; index++) {
            List<?> o = data2.get(index);
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), index, RowType.BODY);
            Row valueRow = this.context.getSheet().createRow(this.context.getSheet().getPhysicalNumberOfRows());
            if (this.context.getBodyHeight() > 0) {
                valueRow.setHeight(this.context.getBodyHeight());
            }
            for (int colIndex = 0, headSize = this.context.getFieldProperties().size(); colIndex < headSize; colIndex++) {
                Object value = o.get(colIndex);
                ExcelFieldProperty property = this.context.getFieldProperties().get(colIndex);
                Cell valueCell = valueRow.createCell(valueRow.getPhysicalNumberOfCells());
                ListenerChain.doSetBodyStyle(this.context.getListenerCache(), valueRow, valueCell, property, index, colIndex);
                value = ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell,
                        property, index, valueCell.getColumnIndex(), RowType.BODY, value);
                ExcelUtils.setCellValue(valueCell, value);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property
                        , index, valueCell.getColumnIndex(), RowType.BODY);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), valueRow, o, index, RowType.BODY);
        }
    }
}
