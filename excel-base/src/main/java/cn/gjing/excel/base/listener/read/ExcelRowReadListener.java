package cn.gjing.excel.base.listener.read;

import cn.gjing.excel.base.meta.RowType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Row listener, which is triggered when the Excel import executor performs Row-related operations
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelRowReadListener<R> extends ExcelReadListener {
    /**
     * Triggered when a row of data has been successfully read
     *
     * @param r        The current row generates object that have value only in binding mode; simple mode does not generate object
     * @param row      The current row
     * @param rowIndex The current row index
     * @param rowType  Current row type
     */
    void readRow(R r, Row row, int rowIndex, RowType rowType);

    /**
     * Triggered when a cell has been successfully read after the data converter has processed it
     *
     * @param cellValue Current cell value
     * @param cell      Current cell
     * @param rowIndex  Current row index
     * @param colIndex  Current col index
     * @param rowType   Current row type
     * @return The contents of the current cell, which you can customize to handle and return
     */
    default Object readCell(Object cellValue, Cell cell, int rowIndex, int colIndex, RowType rowType) {
        return cellValue;
    }

    /**
     * Triggered when all data has been read
     */
    default void readFinish() {
    }

    /**
     * Trigger before reading data
     */
    default void readBefore() {
    }

    /**
     * Continue read next row , once set to false will immediately stop reading down,
     * Triggered after each row is read {@link #readRow}
     *
     * @return true is continue read
     */
    default boolean continueRead() {
        return true;
    }
}
