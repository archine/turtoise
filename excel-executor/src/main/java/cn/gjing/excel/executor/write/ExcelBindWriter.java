package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelClass;
import cn.gjing.excel.base.aware.ExcelWorkbookAware;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.read.ExcelBindReader;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.executor.write.aware.ExcelWriteContextAware;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Excel bind mode writer.
 * The writer needs a mapping entity to correspond to it
 *
 * @author Gjing
 **/
public final class ExcelBindWriter extends ExcelBaseWriter {

    public ExcelBindWriter(ExcelWriterContext context, ExcelClass excelClass, HttpServletResponse response) {
        super(context, excelClass.windowSize(), response, ExecMode.BIND);
    }

    /**
     * To write
     *
     * @param data data
     * @return this
     */
    public ExcelBindWriter write(List<?> data) {
        return this.write(data, this.defaultSheetName, true, null);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName) {
        return this.write(data, sheetName, true, null);
    }

    /**
     * To write
     *
     * @param data     data
     * @param needHead Whether you need excel head
     * @return this
     */
    public ExcelBindWriter write(List<?> data, boolean needHead) {
        return this.write(data, this.defaultSheetName, needHead, null);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @param needHead  Whether you need excel head
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName, boolean needHead) {
        return this.write(data, sheetName, needHead, null);
    }

    /**
     * To write
     *
     * @param data      data
     * @param boxValues dropdown box values
     * @return this
     */
    public ExcelBindWriter write(List<?> data, Map<String, String[]> boxValues) {
        return this.write(data, this.defaultSheetName, true, boxValues);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @param boxValues dropdown box values
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName, Map<String, String[]> boxValues) {
        return this.write(data, sheetName, true, boxValues);
    }

    /**
     * To write
     *
     * @param data      data
     * @param boxValues dropdown box values
     * @param needHead  Whether you need excel head
     * @return this
     */
    public ExcelBindWriter write(List<?> data, boolean needHead, Map<String, String[]> boxValues) {
        return this.write(data, this.defaultSheetName, needHead, boxValues);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @param boxValues dropdown box values
     * @param needHead  Whether you need excel head
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName, boolean needHead, Map<String, String[]> boxValues) {
        this.createSheet(sheetName);
        if (data == null) {
            this.context.setTemplate(true);
            this.writerResolver.writeHead(needHead, boxValues);
        } else {
            this.writerResolver.writeHead(needHead, boxValues)
                    .write(data);
        }
        return this;
    }

    /**
     * Write an Excel header that does not trigger a row callback or cell callback
     *
     * @param bigTitle Big title
     * @return this
     */
    public ExcelBindWriter writeTitle(BigTitle bigTitle) {
        return this.writeTitle(bigTitle, this.defaultSheetName);
    }

    /**
     * Write an Excel header that does not trigger a row listener or cell listener
     *
     * @param bigTitle  Big title
     * @param sheetName Sheet name
     * @return this
     */
    public ExcelBindWriter writeTitle(BigTitle bigTitle, String sheetName) {
        if (bigTitle != null) {
            this.createSheet(sheetName);
            this.writerResolver.writeTitle(bigTitle);
        }
        return this;
    }


    /**
     * Reset Excel mapped entity, Excel file name and file type are not reset
     *
     * @param excelEntity Excel entity
     * @param ignores     The exported field is to be ignored
     * @return this
     */
    public ExcelBindWriter resetExcelClass(Class<?> excelEntity, String... ignores) {
        ExcelClass excel = excelEntity.getAnnotation(ExcelClass.class);
        Objects.requireNonNull(excel, "Failed to reset Excel class, the @Excel annotation was not found on the " + excelEntity);
        List<ExcelFieldProperty> properties = new ArrayList<>();
        this.context.setExcelFields(BeanUtils.getExcelFields(excelEntity, ignores, properties));
        this.context.setExcelEntity(excelEntity);
        this.context.setFieldProperties(properties);
        this.context.setBodyHeight(excel.bodyHeight());
        this.context.setHeaderHeight(excel.headerHeight());
        this.context.setHeaderSeries(properties.get(0).getValue().length);
        return this;
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     *
     * @param enable Whether enable bind, default true
     * @return this
     * @deprecated Please use {@link #bind(String)}
     */
    @Deprecated
    public ExcelBindWriter bind(boolean enable) {
        this.context.setBind(enable);
        return this;
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     * Can be used to {@link ExcelBindReader#check} for a match with an entity class when a file is imported.
     *
     * @param key Unique key ,Each exported file recommends that the key be set to be unique.
     *            Priority is higher than at {@link ExcelClass#uniqueKey()}.
     *            If empty, the unique key in the annotation is used
     * @return this
     */
    public ExcelBindWriter bind(String key) {
        if (StringUtils.hasLength(key)) {
            this.context.setUniqueKey(key);
        }
        this.context.setBind(true);
        return this;
    }

    /**
     * Unbind the unique key of the file
     *
     * @return this
     */
    public ExcelBindWriter unbind() {
        this.context.setBind(false);
        return this;
    }

    /**
     * Add write listener
     *
     * @param listener Write listener
     * @return this
     */
    public ExcelBindWriter addListener(ExcelWriteListener listener) {
        this.context.addListener(listener);
        if (listener instanceof ExcelWriteContextAware) {
            ((ExcelWriteContextAware) listener).setContext(this.context);
        }
        if (listener instanceof ExcelWorkbookAware) {
            ((ExcelWorkbookAware) listener).setWorkbook(this.context.getWorkbook());
        }
        return this;
    }

    /**
     * Add write listeners
     *
     * @param listeners Write listener list
     * @return this
     */
    public ExcelBindWriter addListener(List<? extends ExcelWriteListener> listeners) {
        if (listeners != null) {
            listeners.forEach(this::addListener);
        }
        return this;
    }
}
