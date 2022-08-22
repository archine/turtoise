package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.meta.WRMode;
import cn.gjing.excel.executor.read.ExcelClassReader;
import cn.gjing.excel.executor.util.BeanUtils;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiFunction;

/**
 * Excel simple writer
 * Excel header can be bound with any class.
 *
 * @author Gjing
 **/
public final class ExcelAnyClassWriter extends ExcelBaseWriter {
    /**
     * Field selector
     * assigns a specified field to an Excel field property
     */
    private BiFunction<Integer, List<Field>, Field> fieldSelector;

    public ExcelAnyClassWriter(ExcelWriterContext context, int windowSize, HttpServletResponse response) {
        super(context, windowSize, response, ExecMode.W_ANY_CLASS);
    }

    /**
     * Set the Excel single-level header
     * The order attribute of the generated header field is set to the order in which the elements appear in the header array you pass in, starting at 0
     *
     * @param headNames Excel single-level array of header names
     * @return this
     */
    public ExcelAnyClassWriter head(String... headNames) {
        if (headNames == null || headNames.length == 0) {
            throw new ExcelException("excel header names cannot be null");
        }
        List<ExcelFieldProperty> properties = new ArrayList<>(headNames.length);
        for (String headName : headNames) {
            properties.add(ExcelFieldProperty.builder()
                    .value(new String[]{headName})
                    .build());
        }
        super.context.setHeaderSeries(1);
        super.context.setFieldProperties(properties);
        return this;
    }

    /**
     * Set the Excel header
     * The order attribute of the generated header field is set to the order in which the elements appear in the header array you pass in, starting at 0
     *
     * @param headNames Excel header name arrays, According to the first header array
     *                  size to determine the header hierarchy,
     *                  the subsequent header array must be the same size as the first
     * @return this
     */
    public ExcelAnyClassWriter head(List<String[]> headNames) {
        if (headNames == null || headNames.isEmpty()) {
            throw new ExcelException("excel header names cannot be null");
        }
        List<ExcelFieldProperty> properties = new ArrayList<>(headNames.size());
        for (String[] headName : headNames) {
            properties.add(ExcelFieldProperty.builder()
                    .value(headName)
                    .build());
        }
        super.context.setHeaderSeries(headNames.get(0).length);
        super.context.setFieldProperties(properties);
        return this;
    }

    /**
     * Set the Excel property
     *
     * @param properties Excel filed property, the ExcelFieldProperty order attribute needs to be configured if it needs to be used in listeners
     * @return this
     */
    public ExcelAnyClassWriter head2(List<ExcelFieldProperty> properties) {
        if (properties == null || properties.isEmpty()) {
            throw new ExcelException("excel filed property cannot be null");
        }
        super.context.setFieldProperties(properties);
        super.context.setHeaderSeries(properties.get(0).getValue().length);
        return this;
    }

    /**
     * Set excel head row height
     *
     * @param rowHeight Row height
     * @return this
     */
    public ExcelAnyClassWriter headHeight(short rowHeight) {
        super.context.setHeaderHeight(rowHeight);
        return this;
    }

    /**
     * Set excel body row height
     *
     * @param rowHeight Row height
     * @return this
     */
    public ExcelAnyClassWriter bodyHeight(short rowHeight) {
        super.context.setBodyHeight(rowHeight);
        return this;
    }

    /**
     * To write big title
     *
     * @param bigTitle Big title
     * @return this
     */
    public ExcelAnyClassWriter writeTitle(BigTitle bigTitle) {
        return this.writeTitle(bigTitle, super.defaultSheetName);
    }

    /**
     * To write big title
     *
     * @param bigTitle  Big title
     * @param sheetName Sheet name
     * @return this
     */
    public ExcelAnyClassWriter writeTitle(BigTitle bigTitle, String sheetName) {
        if (bigTitle != null) {
            super.createSheet(sheetName);
            super.writeExecutor.writeTitle(bigTitle);
        }
        return this;
    }

    /**
     * To write
     *
     * @param data Sequential padding, which needs to correspond to the header sequence
     * @return this
     */
    public ExcelAnyClassWriter write(List<?> data) {
        return this.write(data, super.defaultSheetName, true);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @return this
     */
    public ExcelAnyClassWriter write(List<?> data, String sheetName) {
        return this.write(data, sheetName, true);
    }

    /**
     * To write
     *
     * @param data     data
     * @param needHead need to write the header
     * @return this
     */
    public ExcelAnyClassWriter write(List<?> data, boolean needHead) {
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
    public ExcelAnyClassWriter write(List<?> data, String sheetName, boolean needHead) {
        super.createSheet(sheetName);
        if (needHead) {
            super.writeExecutor.writeHead();
        }
        if (data != null && !data.isEmpty()) {
            List<Field> fields = BeanUtils.getAllFields(data.get(0).getClass());
            for (int i = 0, count = super.context.getFieldProperties().size(); i < count; i++) {
                if (this.fieldSelector == null) {
                    super.context.getFieldProperties().get(i).setField(fields.get(i));
                } else {
                    super.context.getFieldProperties().get(i).setField(this.fieldSelector.apply(i, fields));
                }
            }
            super.writeExecutor.writeBody(data);
        }
        return this;
    }

    /**
     * Field selector that assigns a specified field to an Excel field property
     *
     * @param fieldSelector The first parameter is the current Excel field property index（base 0）
     *                      The second parameter is all fields
     *                      The third parameter is the field that you return
     * @return this
     */
    public ExcelAnyClassWriter fieldSelector(BiFunction<Integer, List<Field>, Field> fieldSelector) {
        this.fieldSelector = fieldSelector;
        return this;
    }

    /**
     * Add write listener
     *
     * @param listener Write listener
     * @return this
     */
    public ExcelAnyClassWriter listener(ExcelWriteListener listener) {
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
    public ExcelAnyClassWriter listener(List<? extends ExcelWriteListener> listeners) {
        if (listeners != null) {
            listeners.forEach(this::listener);
        }
        return this;
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     * Can be used to {@link ExcelClassReader#check} for a match with an entity class when a file is imported.
     *
     * @param key Unique key ,Each exported file recommends that the key be set to be unique.
     *            If empty, the binding is invalid
     * @return this
     */
    public ExcelAnyClassWriter bind(String key) {
        if (!StringUtils.hasText(key)) {
            throw new ExcelException("Unique key cannot be empty");
        }
        super.context.setUniqueKey(key);
        super.context.setBind(true);
        return this;
    }

    /**
     * Unbind the unique key of the file
     *
     * @return this
     */
    public ExcelAnyClassWriter unbind() {
        super.context.setBind(false);
        return this;
    }

    /**
     * Set excel write mode
     *
     * @param mode WRMode
     */
    public ExcelAnyClassWriter mode(WRMode mode) {
        super.context.setWrMode(mode);
        return this;
    }
}
