package cn.gjing.excel.base.listener.write;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.meta.RowType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Cell listener, which is triggered when the Excel export executor performs Cell-related operations
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelCellWriteListener extends ExcelWriteListener {
    /**
     * Triggered when the current cell has been written out and is about to start writing the next cell
     *
     * @param sheet         Current sheet
     * @param row           Current row
     * @param cell          Current cell
     * @param dataIndex     Data indexing, depending on the row type, starts at 0
     * @param rowType       Current row type
     * @param property      excel filed property
     */
    void completeCell(Sheet sheet, Row row, Cell cell, ExcelFieldProperty property, int dataIndex, RowType rowType);

    /**
     * Triggered when the data converter finishes processing and is ready to write to the cell
     *
     * @param sheet         Current sheet
     * @param row           Current row
     * @param cell          Current cell
     * @param dataIndex     Data indexing, depending on the row type, starts at 0
     * @param rowType       Current row type
     * @param property      Excel field property
     * @param value         Cell value
     * @return Cell value, if null, no assignment will take place
     */
    default Object assignmentBefore(Sheet sheet, Row row, Cell cell, ExcelFieldProperty property, int dataIndex, RowType rowType, Object value) {
        return value;
    }
}
