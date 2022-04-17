package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ListenerChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Excel simple mode import executor
 *
 * @author Gjing
 **/
public class ExcelSimpleReadExecutor<R> extends ExcelBaseReadExecutor<R> {
    public ExcelSimpleReadExecutor(ExcelReaderContext<R> context) {
        super(context);
    }

    @Override
    public void read(int headerIndex, String sheetName) {
        super.validTemplate();
        super.checkSheet(sheetName);
        ListenerChain.doReadBefore(super.context.getListenerCache());
        boolean continueRead = true;
        for (Row row : super.context.getSheet()) {
            if (!continueRead) {
                break;
            }
            if (row.getRowNum() < headerIndex) {
                continueRead = super.readOther(super.context.getListenerCache(), row);
                continue;
            }
            if (row.getRowNum() == headerIndex) {
                continueRead = super.readHeader(super.context.getListenerCache(), row);
                continue;
            }
            for (int i = 0, size = super.context.getHeadNames().size(); i < size; i++) {
                String head = super.context.getHeadNames().get(i);
                if ("ignored".equals(head)) {
                    continue;
                }
                Cell cell = row.getCell(i + super.startCol);
                Object value;
                if (cell != null) {
                    value = this.getValue(null, cell, null, false, false, RowType.BODY, ExecMode.SIMPLE_READ);
                    ListenerChain.doReadCell(super.context.getListenerCache(), value, cell, row.getRowNum(), i, RowType.BODY);
                } else {
                    ListenerChain.doReadCell(super.context.getListenerCache(), null, null, row.getRowNum(), i, RowType.BODY);
                }
            }
            continueRead = ListenerChain.doReadRow(super.context.getListenerCache(), null, row, RowType.BODY);
        }
        ListenerChain.doReadFinish(super.context.getListenerCache());
    }
}
