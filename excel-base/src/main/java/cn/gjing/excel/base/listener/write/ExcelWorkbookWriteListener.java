package cn.gjing.excel.base.listener.write;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * WorkBook listener, which is triggered when the Excel export executor performs workbook-related operations
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelWorkbookWriteListener extends ExcelWriteListener {
    /**
     * Refresh WorkBook data to the front of the response flow
     *
     * @param workbook Current workbook
     * @return If false, the flush is aborted, meaning that the data will not be exported
     */
    boolean flushBefore(Workbook workbook);
}
