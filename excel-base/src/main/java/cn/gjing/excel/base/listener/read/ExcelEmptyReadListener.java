package cn.gjing.excel.base.listener.read;

import cn.gjing.excel.base.annotation.ExcelField;

/**
 * Body cell null-value listener, triggered when Excel import execution detects that cell content is not present
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelEmptyReadListener<R> extends ExcelReadListener {
    /**
     * When a body cell is read, if the cell does not exist or the value is empty,
     * and header is set as required in the mapping entity {@link ExcelField#required()}.
     *
     * Return true to continue reading the cells of that row and retain the object generated by the current row,
     * returning false immediately stops reading the current row and starts the next row,
     * and deletes the objects generated by the current row,
     * you can also abort the import by throwing an exception
     *
     * @param r        Current Java object
     * @param rowIndex Current row index
     * @param colIndex Current column index
     * @return Whether to continue reading the row
     */
    boolean readEmpty(R r, int rowIndex, int colIndex);
}
