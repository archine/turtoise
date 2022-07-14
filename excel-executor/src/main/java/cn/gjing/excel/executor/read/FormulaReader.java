package cn.gjing.excel.executor.read;

import cn.gjing.excel.base.meta.RowType;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;

/**
 * Formula type cell data reader
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface FormulaReader {
    /**
     * Read formula
     *
     * @param cell    Current cell
     * @param field   Current filed
     * @param rowType Current row type
     * @return cell value
     */
    Object read(Cell cell, Field field, RowType rowType);
}
