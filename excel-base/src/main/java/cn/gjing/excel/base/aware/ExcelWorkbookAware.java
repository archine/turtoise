package cn.gjing.excel.base.aware;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Workbook loader, through which you can obtain the Workbook
 *
 * @author Gjing
 **/
public interface ExcelWorkbookAware extends ExcelAware {
    /**
     * Set workbook
     *
     * @param workbook workbook
     */
    void setWorkbook(Workbook workbook);
}
