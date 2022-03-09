package cn.gjing.excel.base.listener.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Excel style listener
 *
 * @author Gjing
 **/
public interface ExcelStyleWriteListener extends ExcelWriteListener {
    /**
     * Set excel big title style
     *
     * @param cell     Current cell
     * @param bigTitle Big title
     */
    void setTitleStyle(BigTitle bigTitle, Cell cell);

    /**
     * Setting excel header cell styles is triggered when each header cell is successfully created
     *
     * @param row       Current row
     * @param cell      Current cell
     * @param dataIndex Data index, starts at 0
     * @param colIndex  cell index
     * @param property  ExcelField property of current field
     */
    void setHeadStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex);

    /**
     * Setting excel body cell styles is triggered when each body cell is successfully created
     *
     * @param row       Current row
     * @param cell      Current cell
     * @param dataIndex Data index, starts at 0
     * @param colIndex  cell index
     * @param property  ExcelField property of current field
     */
    void setBodyStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex);
}
