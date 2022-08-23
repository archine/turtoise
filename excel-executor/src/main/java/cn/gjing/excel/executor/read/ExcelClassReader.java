package cn.gjing.excel.executor.read;

import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.read.ExcelReadListener;
import cn.gjing.excel.base.listener.read.ExcelResultReadListener;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.meta.WRMode;
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
public final class ExcelClassReader<R> extends ExcelBaseReader<R> {
    public ExcelClassReader(ExcelReaderContext<R> context, InputStream inputStream, ExcelType excelType, int cacheRowSize, int bufferSize) {
        super(context, inputStream, excelType, cacheRowSize, bufferSize, ExecMode.R_Class);
    }

    /**
     * Read excel
     * By default, the index of the first row of Sheet is used as the index of the table head
     *
     * @return this
     */
    public ExcelClassReader<R> read() {
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
    public ExcelClassReader<R> read(String sheetName) {
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
    public ExcelClassReader<R> read(int headerIndex) {
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
    public ExcelClassReader<R> read(int headerIndex, String sheetName) {
        super.baseReadExecutor.read(headerIndex, sheetName);
        return this;
    }

    /**
     * Whether to read all rows before the header
     *
     * @param need Need
     * @return this
     */
    public ExcelClassReader<R> readOther(boolean need) {
        super.context.setReadOther(need);
        return this;
    }

    /**
     * Check whether the imported Excel file matches the Excel mapping entity class.
     * Thrown {@link ExcelTemplateException} if there is don't match.
     *
     * @return this
     **/
    public ExcelClassReader<R> check() {
        super.context.setCheckTemplate(true);
        return this;
    }

    /**
     * Check whether the imported Excel file matches the Excel mapping entity class.
     * Thrown {@link ExcelTemplateException} if is don't match.
     *
     * @param idCard Excel file id card
     * @return this
     **/
    public ExcelClassReader<R> check(String idCard) {
        if (!StringUtils.hasText(idCard)) {
            throw new ExcelException("idCard cannot be empty");
        }
        super.context.setCheckTemplate(true);
        super.context.setIdCard(idCard);
        return this;
    }

    /**
     * Add excel read listener
     *
     * @param readListenerList Read listeners
     * @return this
     */
    public ExcelClassReader<R> listener(List<? extends ExcelReadListener> readListenerList) {
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
    public ExcelClassReader<R> listener(ExcelReadListener readListener) {
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
    public ExcelClassReader<R> subscribe(ExcelResultReadListener<R> excelResultReadListener) {
        super.context.setResultReadListener(excelResultReadListener);
        return this;
    }

    /**
     * Set excel import mode
     *
     * @param mode WRMode
     */
    public ExcelClassReader<R> mode(WRMode mode) {
        super.context.setWrMode(mode);
        return this;
    }
}