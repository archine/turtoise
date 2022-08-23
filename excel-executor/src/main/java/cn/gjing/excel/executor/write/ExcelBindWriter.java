package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.ExcelListener;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.meta.WRMode;
import cn.gjing.excel.executor.read.ExcelClassReader;
import cn.gjing.excel.executor.util.BeanUtils;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Objects;
import java.util.function.Predicate;

/**
 * Excel class writer
 * The Excel header is bound to the header field of the current class.
 *
 * @author Gjing
 **/
public final class ExcelBindWriter extends ExcelBaseWriter {

    public ExcelBindWriter(ExcelWriterContext context, Excel excel, HttpServletResponse response) {
        super(context, excel.windowSize(), response, ExecMode.W_FIXED_CLASS);
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
        if (needHead) {
            super.writeExecutor.writeHead();
        }
        if (data != null && !data.isEmpty()) {
            super.writeExecutor.writeBody(data);
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
            super.writeExecutor.writeTitle(bigTitle);
        }
        return this;
    }

    /**
     * Reset Excel entity, file names and unique keys (if present) do not change
     *
     * @param excelEntity Excel entity
     * @param ignores     The exported field is to be ignored
     * @return this
     */
    public ExcelBindWriter resetEntity(Class<?> excelEntity, String... ignores) {
        Excel excel = excelEntity.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "Failed to reset Excel class, the @Excel annotation was not found on the " + excelEntity);
        super.context.setFieldProperties(BeanUtils.getExcelFiledProperties(excelEntity, ignores));
        super.context.setExcelEntity(excelEntity);
        super.context.setBodyHeight(excel.bodyHeight());
        super.context.setHeaderHeight(excel.headerHeight());
        super.context.setHeaderSeries(super.context.getFieldProperties().size() == 0 ? 0 : super.context.getFieldProperties().get(0).getValue().length);
        return this;
    }

    /**
     * Clears listeners in the current context
     *
     * @param predicate The assertion condition, true, is removed
     * @return this
     */
    public ExcelBindWriter resetListeners(Predicate<ExcelListener> predicate) {
        super.context.getListenerCache().removeIf(predicate);
        return this;
    }

    /**
     * Clears all listeners in the current context, which triggers the listener initializer again
     * <p>
     * Excel entities in the current context are passed to the listener initializer, so they should be called after the resetEntity method
     *
     * @return this
     */
    public ExcelBindWriter resetListeners() {
        return this.resetListeners((l) -> true);
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     * Can be used to {@link ExcelClassReader#check} for a match with an entity class when a file is imported.
     *
     * @param idCard Excel file id card ,Each exported file recommends that the key be set to be unique.
     *               Priority is higher than at {@link Excel#idCard()}}.
     *               If empty, the unique key in the annotation is used
     * @return this
     */
    public ExcelBindWriter bind(String idCard) {
        if (!StringUtils.hasText(idCard)) {
            throw new ExcelException("id card cannot be empty");
        }
        super.context.setIdCard(idCard);
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
    public ExcelBindWriter listener(ExcelWriteListener listener) {
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
    public ExcelBindWriter listener(List<? extends ExcelWriteListener> listeners) {
        if (listeners != null) {
            listeners.forEach(this::listener);
        }
        return this;
    }

    /**
     * Set excel write mode
     *
     * @param mode WRMode
     */
    public ExcelBindWriter mode(WRMode mode) {
        super.context.setWrMode(mode);
        return this;
    }
}
