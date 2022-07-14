package cn.gjing.excel.executor.read;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.read.ExcelReadListener;
import cn.gjing.excel.base.listener.read.ExcelResultReadListener;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import org.springframework.util.StringUtils;

import java.io.InputStream;
import java.util.List;

/**
 * Excel bind mode reader
 * The reader needs a mapping entity to correspond to it,
 * Automatically convert the data for each row to the corresponding Java entity
 *
 * @author Gjing
 **/
public final class ExcelBindReader<R> extends ExcelBaseReader<R> {
    public ExcelBindReader(ExcelReaderContext<R> context, InputStream inputStream, ExcelType excelType, int cacheRowSize, int bufferSize) {
        super(context, inputStream, excelType, cacheRowSize, bufferSize, ExecMode.BIND_READ);
    }

    /**
     * Read excel
     * By default, the index of the first row of Sheet is used as the index of the table head
     *
     * @return this
     */
    public ExcelBindReader<R> read() {
        super.baseReadExecutor.read(0, this.defaultSheetName);
        return this;
    }

    /**
     * Read the specified sheet
     * By default, the index of the first row of Sheet is used as the index of the table head
     *
     * @param sheetName sheet name
     * @return this
     */
    public ExcelBindReader<R> read(String sheetName) {
        super.baseReadExecutor.read(0, sheetName);
        return this;
    }

    /**
     * Specifies that the Excel subscript to start reading.
     * This line must be a real subscript,
     *
     * @param headerIndex The subscript of the table header. If there are multiple levels of table headers,
     *                    set the subscript of the bottom level of the table header. The index starts at 0
     * @return this
     */
    public ExcelBindReader<R> read(int headerIndex) {
        super.baseReadExecutor.read(headerIndex, this.defaultSheetName);
        return this;
    }

    /**
     * Read the specified sheet
     *
     * @param headerIndex The subscript of the table header. If there are multiple levels of table headers,
     *                    set the subscript of the bottom level of the table header. The index starts at 0
     * @param sheetName   Excel Sheet name
     * @return this
     */
    public ExcelBindReader<R> read(int headerIndex, String sheetName) {
        super.baseReadExecutor.read(headerIndex, sheetName);
        return this;
    }

    /**
     * Whether to read all rows before the header
     *
     * @param need Need
     * @return this
     */
    public ExcelBindReader<R> readOther(boolean need) {
        super.context.setReadOther(need);
        return this;
    }

    /**
     * Check whether the imported Excel file matches the Excel mapping entity class.
     * Thrown {@link ExcelTemplateException} if there is don't match.
     *
     * @return this
     **/
    public ExcelBindReader<R> check() {
        super.context.setCheckTemplate(true);
        return this;
    }

    /**
     * Check whether the imported Excel file matches the Excel mapping entity class.
     * Thrown {@link ExcelTemplateException} if is don't match.
     *
     * @param key Unique key
     * @return this
     **/
    public ExcelBindReader<R> check(String key) {
        if (!StringUtils.hasText(key)) {
            throw new ExcelException("Unique key cannot be empty");
        }
        super.context.setCheckTemplate(true);
        super.context.setUniqueKey(key);
        return this;
    }

    /**
     * Add excel read listener
     *
     * @param readListenerList Read listeners
     * @return this
     */
    public ExcelBindReader<R> listener(List<? extends ExcelReadListener> readListenerList) {
        if (readListenerList != null) {
            readListenerList.forEach(this::listener);
        }
        return this;
    }

    /**
     * Add excel read listener
     *
     * @param readListener Read listener
     * @return this
     */
    public ExcelBindReader<R> listener(ExcelReadListener readListener) {
        super.context.addListener(readListener);
        super.initAware(readListener);
        return this;
    }

    /**
     * Subscribe to the data after the import is complete
     *
     * @param excelResultReadListener resultReadListener
     * @return this
     */
    public ExcelBindReader<R> subscribe(ExcelResultReadListener<R> excelResultReadListener) {
        super.context.setResultReadListener(excelResultReadListener);
        return this;
    }

    /**
     * Set the current write position
     *
     * @param startCol col index, based on 0
     * @return this
     */
    public ExcelBindReader<R> withPosition(int startCol) {
        if (startCol < 0) {
            throw new ExcelException("the column index to start reading cannot be less than 0");
        }
        super.baseReadExecutor.setPosition(startCol);
        return this;
    }

    /**
     * Set the formula reader
     *
     * @param formulaReader FormulaReader
     * @return this
     */
    public ExcelBindReader<R> setFormulaReader(FormulaReader formulaReader) {
        super.baseReadExecutor.setFormulaReader(formulaReader);
        return this;
    }
}
