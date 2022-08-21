//package cn.gjing.excel.executor.write.core;
//
//import cn.gjing.excel.base.ExcelFieldProperty;
//import cn.gjing.excel.base.context.ExcelWriterContext;
//import cn.gjing.excel.base.meta.RowType;
//import cn.gjing.excel.executor.WRMode;
//import cn.gjing.excel.executor.util.ExcelUtils;
//import cn.gjing.excel.executor.util.ListenerChain;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//
//import java.util.List;
//
///**
// * Export processor for Excel simple mode
// *
// * @author Gjing
// **/
//public class ExcelSimpleWriterExecutor extends ExcelBaseWriteExecutor {
//    public ExcelSimpleWriterExecutor(ExcelWriterContext context) {
//        super(context);
//    }
//
//    @SuppressWarnings("unchecked")
//    @Override
//    public void writeBody(List<?> data) {
//        List<List<Object>> data2 = (List<List<Object>>) data;
////        for (int dataIndex = 0, dataSize = data.size(); dataIndex < dataSize; dataIndex++) {
//            List<?> o = data2.get(dataIndex);
//            ListenerChain.doCreateRowBefore(this.context.getListenerCache(), this.context.getSheet(), dataIndex, RowType.BODY);
//            Row valueRow = this.context.getSheet().createRow(this.context.getSheet().getLastRowNum() + 1);
//            if (this.context.getBodyHeight() > 0) {
//                valueRow.setHeight(this.context.getBodyHeight());
//            }
//            for (int headerIndex = 0, headSize = this.context.getFieldProperties().size(); headerIndex < headSize; headerIndex++) {
//                Object value = o.get(headerIndex);
//                ExcelFieldProperty property = this.context.getFieldProperties().get(headerIndex);
//                short lastCellNum = super.wrMode == WRMode.INDEX ? property.getIndex() : valueRow.getLastCellNum();
//                Cell valueCell = valueRow.createCell(lastCellNum == -1 ? 0 : lastCellNum);
//                ListenerChain.doSetBodyStyle(this.context.getListenerCache(), valueRow, valueCell, property, dataIndex);
//                value = ListenerChain.doAssignmentBefore(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell,
//                        property, dataIndex, RowType.BODY, value);
//                ExcelUtils.setCellValue(valueCell, value);
//                ListenerChain.doCompleteCell(this.context.getListenerCache(), this.context.getSheet(), valueRow, valueCell, property,
//                        dataIndex, RowType.BODY);
//            }
//            ListenerChain.doCompleteRow(this.context.getListenerCache(), this.context.getSheet(), valueRow, o, dataIndex, RowType.BODY);
//        }
//    }
//}
