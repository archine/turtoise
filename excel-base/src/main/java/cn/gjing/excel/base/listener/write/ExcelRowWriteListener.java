package cn.gjing.excel.base.listener.write;

import cn.gjing.excel.base.meta.RowType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Row listener, which is triggered when the Excel export executor performs Row-related operations
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelRowWriteListener extends ExcelWriteListener {
    /**
     * Triggered when all header fields of the Excel entity corresponding to the current row have been written out
     *
     * @param sheet       Current sheet
     * @param row         Current row
     * @param excelEntity Excel entity corresponding to the current row
     * @param dataIndex   Data index, depending on the row type, starts at 0
     * @param rowType     Current row type
     */
    void completeRow(Sheet sheet, Row row, Object excelEntity, int dataIndex, RowType rowType);

    /**
     * Triggered before a new row is created
     *
     * @param sheet     Current sheet
     * @param dataIndex New row index, depending on the row type, starts at 0
     * @param rowType   New row type
     */
    default void createBefore(Sheet sheet, int dataIndex, RowType rowType) {

    }
}
