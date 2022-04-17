package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.base.util.ListenerChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * Export processor for Excel simple mode
 *
 * @author Gjing
 **/
public class ExcelSimpleWriterExecutor extends ExcelBaseWriteExecutor {
    public ExcelSimpleWriterExecutor(ExcelWriterContext context) {
        super(context);
    }

    @Override
    public void writeHead() {
        Row headRow;
        for (int index = 0; index < this.context.getHeaderSeries(); index++) {
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), index, RowType.HEAD);
            headRow = this.context.getSheet().createRow(this.context.getSheet().getLastRowNum() + 1);
            if (this.context.getHeaderHeight() > 0) {
                headRow.setHeight(this.context.getHeaderHeight());
            }
            for (int headerIndex = 0, headSize = this.context.getFieldProperties().size(); headerIndex < headSize; headerIndex++) {
                String headName = this.context.getFieldProperties().get(headerIndex).getValue()[index];
                ExcelFieldProperty property = this.context.getFieldProperties().get(headerIndex);
                short lastCellNum = headRow.getLastCellNum();
                Cell headCell = headRow.createCell(lastCellNum == -1 ? super.startCol : lastCellNum);
                ListenerChain.doSetHeadStyle(this.context.getListenerCache(), headRow, headCell, property, index);
                headName = (String) ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(),
                        headRow, headCell, property, index, RowType.HEAD, headName);
                headCell.setCellValue(headName);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), headRow, headCell, property,
                        index, RowType.HEAD);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), headRow, null, index, RowType.HEAD);
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public void writeBody(List<?> data) {
        List<List<Object>> data2 = (List<List<Object>>) data;
        for (int index = 0, dataSize = data.size(); index < dataSize; index++) {
            List<?> o = data2.get(index);
            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), index, RowType.BODY);
            Row valueRow = this.context.getSheet().createRow(this.context.getSheet().getLastRowNum() + 1);
            if (this.context.getBodyHeight() > 0) {
                valueRow.setHeight(this.context.getBodyHeight());
            }
            for (int headerIndex = 0, headSize = this.context.getFieldProperties().size(); headerIndex < headSize; headerIndex++) {
                Object value = o.get(headerIndex);
                ExcelFieldProperty property = this.context.getFieldProperties().get(headerIndex);
                short lastCellNum = valueRow.getLastCellNum();
                Cell valueCell = valueRow.createCell(lastCellNum == -1 ? super.startCol : lastCellNum);
                ListenerChain.doSetBodyStyle(this.context.getListenerCache(), valueRow, valueCell, property, index);
                value = ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell,
                        property, index, RowType.BODY, value);
                ExcelUtils.setCellValue(valueCell, value);
                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property,
                        index, RowType.BODY);
            }
            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), valueRow, o, index, RowType.BODY);
        }
    }
}
