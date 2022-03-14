package cn.gjing.excel.base.listener.write;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Sheet listener, which is triggered when the Excel export executor performs Sheet-related operations
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelSheetWriteListener extends ExcelWriteListener {
    /**
     * Trigger when Sheet is created
     *
     * @param sheet Current created sheet
     */
    void completeSheet(Sheet sheet);
}
