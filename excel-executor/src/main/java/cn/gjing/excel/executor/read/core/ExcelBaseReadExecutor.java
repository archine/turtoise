package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.base.util.ParamUtils;
import cn.gjing.excel.executor.util.ListenerChain;
import com.monitorjbl.xlsx.exceptions.MissingSheetException;
import com.monitorjbl.xlsx.impl.StreamingWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel base reader executor
 *
 * @author Gjing
 **/
public abstract class ExcelBaseReadExecutor<R> {
    protected final ExcelReaderContext<R> context;
    protected boolean saveCurrentRowObj;

    public ExcelBaseReadExecutor(ExcelReaderContext<R> context) {
        this.context = context;
    }

    /**
     * Start Import excel
     *
     * @param headerIndex Excel header index
     * @param sheetName   sheetName
     */
    public abstract void read(int headerIndex, String sheetName);

    /**
     * Read head row
     *
     * @param row              Current row
     * @return Continue read next row
     */
    protected boolean readHeader(Row row) {
        for (Cell cell : row) {
            Object value = this.getValue(null, cell, true, false);
            this.context.getHeadNames().add(ListenerChain.doReadCell(this.context.getListenerCache(), value, cell, row.getRowNum(), cell.getColumnIndex(), RowType.HEAD));
        }
        return ListenerChain.doReadRow(this.context.getListenerCache(), null, row, RowType.HEAD);
    }

    /**
     * Reads all rows before the table header
     *
     * @param row              Current row
     * @return Continue read next row
     */
    protected boolean readOther(Row row) {
        if (this.context.isReadOther()) {
            for (Cell cell : row) {
                Object value = this.getValue(null, cell, false, false);
                ListenerChain.doReadCell(this.context.getListenerCache(), value, cell, row.getRowNum(), cell.getColumnIndex(), RowType.OTHER);
            }
            return ListenerChain.doReadRow(this.context.getListenerCache(), null, row, RowType.OTHER);
        }
        return true;
    }

    /**
     * Check sheet is exists
     *
     * @param sheetName Sheet name
     */
    protected void checkSheet(String sheetName) {
        if (this.context.getWorkbook() instanceof StreamingWorkbook) {
            try {
                this.context.setSheet(this.context.getWorkbook().getSheet(sheetName));
            } catch (MissingSheetException e) {
                throw new ExcelException("The " + sheetName + " is not found in the workbook");
            }
        } else {
            Sheet sheet = this.context.getWorkbook().getSheet(sheetName);
            if (sheet == null) {
                throw new ExcelException("The " + sheetName + " is not found in the workbook");
            }
            this.context.setSheet(sheet);
        }
    }

    protected void validTemplate() {
        if (this.context.isCheckTemplate()) {
            String key = "excelUnqSheet";
            if (this.context.getWorkbook().getSheetIndex(key) == -1) {
                throw new ExcelTemplateException();
            }
            for (Row row : this.context.getWorkbook().getSheet(key)) {
                if (!ParamUtils.equals(ParamUtils.encodeMd5(this.context.getIdCard()), row.getCell(0).getStringCellValue())) {
                    throw new ExcelTemplateException();
                }
                break;
            }
            this.context.setCheckTemplate(false);
        }
    }

    /**
     * Get the value of the cell
     *
     * @param cell     cell
     * @param trim     Remove white space on both sides of the string
     * @param required Cell content required
     * @param r        Current row generated row
     * @return value
     */
    protected Object getValue(R r, Cell cell, boolean trim, boolean required) {
        Object cellValue = ExcelUtils.getCellValue(cell, cell.getCellType(), trim);
        if (cellValue == null) {
            if (required) {
                this.saveCurrentRowObj = ListenerChain.doReadEmpty(this.context.getListenerCache(), r, cell.getRowIndex(), cell.getColumnIndex());
            }
        }
        return cellValue;
    }
}
