package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.ExcelListener;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ListenerChain;
import cn.gjing.excel.base.util.ParamUtils;
import cn.gjing.excel.executor.read.FormulaReader;
import com.monitorjbl.xlsx.impl.StreamingWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Excel base reader executor
 *
 * @author Gjing
 **/
public abstract class ExcelBaseReadExecutor<R> {
    protected final ExcelReaderContext<R> context;
    protected boolean saveCurrentRowObj;
    protected int startCol;
    protected FormulaReader formulaReader;

    public ExcelBaseReadExecutor(ExcelReaderContext<R> context) {
        this.context = context;
    }

    /**
     * Sets the location where data is to be written
     *
     * @param startCol column index
     */
    public void setPosition(int startCol) {
        this.startCol = startCol;
    }

    /**
     * Set the formula reader
     *
     * @param formulaReader FormulaReader
     */
    public void setFormulaReader(FormulaReader formulaReader) {
        this.formulaReader = formulaReader;
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
     * @param rowReadListeners Row read listener
     * @param row              Current row
     * @return Continue read next row
     */
    protected boolean readHeader(List<ExcelListener> rowReadListeners, Row row) {
        for (Cell cell : row) {
            if (cell.getColumnIndex() < this.startCol) {
                continue;
            }
            String value = cell.getStringCellValue();
            if (ParamUtils.contains(this.context.getIgnores(), value)) {
                value = "ignored";
            }
            this.context.getHeadNames().add(String.valueOf(ListenerChain.doReadCell(rowReadListeners, value, cell, row.getRowNum(), cell.getColumnIndex(), RowType.HEAD)));
        }
        return ListenerChain.doReadRow(rowReadListeners, null, row, RowType.HEAD);
    }

    /**
     * Reads all rows before the table header
     *
     * @param rowReadListeners Row read listener
     * @param row              Current row
     * @return Continue read next row
     */
    protected boolean readOther(List<ExcelListener> rowReadListeners, Row row) {
        if (this.context.isReadOther()) {
            for (Cell cell : row) {
                Object value = this.getValue(null, cell, null, false, false, RowType.OTHER);
                ListenerChain.doReadCell(rowReadListeners, value, cell, row.getRowNum(), cell.getColumnIndex(), RowType.OTHER);
            }
            return ListenerChain.doReadRow(rowReadListeners, null, row, RowType.OTHER);
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
            } catch (Exception e) {
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
                if (!ParamUtils.equals(ParamUtils.encodeMd5(this.context.getUniqueKey()), row.getCell(0).getStringCellValue())) {
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
     * @param field    Current field
     * @param r        Current row generated row
     * @param rowType  rowType Current row type
     * @return value
     */
    protected Object getValue(R r, Cell cell, Field field, boolean trim, boolean required, RowType rowType) {
        switch (cell.getCellType()) {
            case _NONE:
            case BLANK:
            case ERROR:
                if (rowType == RowType.BODY) {
                    if (required) {
                        this.saveCurrentRowObj = ListenerChain.doReadEmpty(this.context.getListenerCache(), r, cell.getRowIndex(), cell.getColumnIndex());
                    }
                }
                break;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case FORMULA:
                if (formulaReader == null) {
                    throw new ExcelException("The current Excel file has cells of formula type, so you need to set the formula reader");
                }
                return this.formulaReader.read(cell, field, rowType);
            default:
                return trim ? cell.getStringCellValue().trim() : cell.getStringCellValue();
        }
        return null;
    }
}
