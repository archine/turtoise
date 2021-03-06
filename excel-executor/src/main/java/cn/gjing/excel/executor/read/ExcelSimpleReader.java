package cn.gjing.excel.executor.read;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.read.ExcelReadListener;
import cn.gjing.excel.base.listener.read.ExcelRowReadListener;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import org.springframework.util.StringUtils;

import java.io.InputStream;
import java.util.List;

/**
 * Excel simple mode reader
 * No mapping entities need to be provided.
 * Instead of automatically turning each row into a Java entity,
 * you can manually assemble your own objects in {@link ExcelRowReadListener}
 *
 * @author Gjing
 **/
public class ExcelSimpleReader<R> extends ExcelBaseReader<R> {
    public ExcelSimpleReader(ExcelReaderContext<R> context, InputStream inputStream, ExcelType excelType, int cacheRowSize, int bufferSize) {
        super(context, inputStream, excelType, cacheRowSize, bufferSize, ExecMode.SIMPLE_READ);
    }

    /**
     * Read excel
     * By default, the index of the first row of Sheet is used as the index of the table head
     *
     * @return this
     */
    public ExcelSimpleReader<R> read() {
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
    public ExcelSimpleReader<R> read(String sheetName) {
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
    public ExcelSimpleReader<R> read(int headerIndex) {
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
    public ExcelSimpleReader<R> read(int headerIndex, String sheetName) {
        super.baseReadExecutor.read(headerIndex, sheetName);
        return this;
    }

    /**
     * Whether to read all rows before the header
     *
     * @param need Need
     * @return this
     */
    public ExcelSimpleReader<R> readOther(boolean need) {
        super.context.setReadOther(need);
        return this;
    }

    /**
     * Check whether the imported Excel file matches the Excel mapping entity class.
     * Thrown {@link ExcelTemplateException} if is don't match.
     *
     * @param key Unique key
     * @return this
     **/
    public ExcelSimpleReader<R> check(String key) {
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
    public ExcelSimpleReader<R> listener(List<? extends ExcelReadListener> readListenerList) {
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
    public ExcelSimpleReader<R> listener(ExcelReadListener readListener) {
        super.context.addListener(readListener);
        super.initAware(readListener);
        return this;
    }

    /**
     * Set the current write position
     *
     * @param startCol col index, based on 0
     * @return this
     */
    public ExcelSimpleReader<R> withPosition(int startCol) {
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
    public ExcelSimpleReader<R> setFormulaReader(FormulaReader formulaReader) {
        super.baseReadExecutor.setFormulaReader(formulaReader);
        return this;
    }
}
