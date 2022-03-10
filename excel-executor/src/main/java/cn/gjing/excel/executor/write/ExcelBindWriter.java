package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.read.ExcelBindReader;
import cn.gjing.excel.executor.util.BeanUtils;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Objects;

/**
 * Excel bind mode writer.
 * The writer needs a mapping entity to correspond to it
 *
 * @author Gjing
 **/
public final class ExcelBindWriter extends ExcelBaseWriter {

    public ExcelBindWriter(ExcelWriterContext context, Excel excel, HttpServletResponse response) {
        super(context, excel.windowSize(), response, ExecMode.BIND);
    }

    /**
     * To write
     *
     * @param data data
     * @return this
     */
    public ExcelBindWriter write(List<?> data) {
        return this.write(data, super.defaultSheetName, true);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName) {
        return this.write(data, sheetName, true);
    }

    /**
     * To write
     *
     * @param data     data
     * @param needHead need to write the header
     * @return this
     */
    public ExcelBindWriter write(List<?> data, boolean needHead) {
        return this.write(data, super.defaultSheetName, needHead);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @param needHead  need to write the header
     * @return this
     */
    public ExcelBindWriter write(List<?> data, String sheetName, boolean needHead) {
        super.createSheet(sheetName);
        if (data == null) {
            super.context.setTemplate(true);
            this.writerResolver.writeHead(needHead);
        } else {
            this.writerResolver.writeHead(needHead).write(data);
        }
        return this;
    }

    /**
     * To write big title
     *
     * @param bigTitle Big title
     * @return this
     */
    public ExcelBindWriter writeTitle(BigTitle bigTitle) {
        return this.writeTitle(bigTitle, super.defaultSheetName);
    }

    /**
     * To write big title
     *
     * @param bigTitle  Big title
     * @param sheetName Sheet name
     * @return this
     */
    public ExcelBindWriter writeTitle(BigTitle bigTitle, String sheetName) {
        if (bigTitle != null) {
            super.createSheet(sheetName);
            super.writerResolver.writeTitle(bigTitle);
        }
        return this;
    }

    /**
     * Reset Excel mapped entity, File names and unique keys (if present) do not change
     *
     * @param excelEntity Excel entity
     * @param ignores     The exported field is to be ignored
     * @return this
     */
    public ExcelBindWriter resetExcelEntity(Class<?> excelEntity, String... ignores) {
        Excel excel = excelEntity.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "Failed to reset Excel class, the @Excel annotation was not found on the " + excelEntity);
        super.context.setFieldProperties(BeanUtils.getExcelFiledProperties(excelEntity, ignores));
        super.context.setExcelEntity(excelEntity);
        super.context.setBodyHeight(excel.bodyHeight());
        super.context.setHeaderHeight(excel.headerHeight());
        super.context.setHeaderSeries(super.context.getFieldProperties().get(0).getValue().length);
        return this;
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     * Can be used to {@link ExcelBindReader#check} for a match with an entity class when a file is imported.
     *
     * @param key Unique key ,Each exported file recommends that the key be set to be unique.
     *            Priority is higher than at {@link Excel#uniqueKey()}.
     *            If empty, the unique key in the annotation is used
     * @return this
     */
    public ExcelBindWriter bind(String key) {
        if (StringUtils.hasLength(key)) {
            super.context.setUniqueKey(key);
        }
        super.context.setBind(true);
        return this;
    }

    /**
     * Unbind the unique key of the file
     *
     * @return this
     */
    public ExcelBindWriter unbind() {
        super.context.setBind(false);
        return this;
    }

    /**
     * Add write listener
     *
     * @param listener Write listener
     * @return this
     */
    public ExcelBindWriter addListener(ExcelWriteListener listener) {
        super.context.addListener(listener);
        super.initAware(listener);
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
