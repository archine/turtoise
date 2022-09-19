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
public final class ExcelSimpleWriter extends ExcelBaseWriter {
    /**
     * Field selector
     * assigns a specified field to an Excel field property
     */
    private BiFunction<Integer, List<Field>, Field> fieldSelector;

    public ExcelSimpleWriter(ExcelWriterContext context, int windowSize, HttpServletResponse response) {
        super(context, windowSize, response, ExecMode.W_SIMPLE);
    }

    /**
     * Set the Excel single-level header
     * The order attribute of the generated header field is set to the order in which the elements appear in the header array you pass in, starting at 0
     *
     * @param headers Excel header properties, supported types are string[], String, ExcelFieldProperties
     * @return this
     */
    public ExcelSimpleWriter head(Object... headers) {
        if (headers == null || headers.length == 0) {
            super.close();
            throw new ExcelException("excel headers cannot be null");
        }
        List<ExcelFieldProperty> properties = new ArrayList<>(headers.length);
        for (Object header : headers) {
            if (header instanceof String) {
                properties.add(ExcelFieldProperty.builder()
                        .value(new String[]{header.toString()})
                        .build());
                continue;
            }
            if (header instanceof ExcelFieldProperty) {
                properties.add((ExcelFieldProperty) header);
                continue;
            }
            super.close();
            throw new IllegalArgumentException("invalid header value,supports ExcelFieldProperty, string");
        }
        super.context.setHeaderSeries(properties.get(0).getValue().length);
        super.context.setFieldProperties(properties);
        return this;
    }

    /**
     * Set excel head row height
     *
     * @param rowHeight Row height
     * @return this
     */
    public ExcelSimpleWriter headHeight(short rowHeight) {
        super.context.setHeaderHeight(rowHeight);
        return this;
    }

    /**
     * Set excel body row height
     *
     * @param rowHeight Row height
     * @return this
     */
    public ExcelSimpleWriter bodyHeight(short rowHeight) {
        super.context.setBodyHeight(rowHeight);
        return this;
    }

    /**
     * To write big title
     *
     * @param bigTitle Big title
     * @return this
     */
    public ExcelSimpleWriter writeTitle(BigTitle bigTitle) {
        return this.writeTitle(bigTitle, super.defaultSheetName);
    }

    /**
     * To write big title
     *
     * @param bigTitle  Big title
     * @param sheetName Sheet name
     * @return this
     */
    public ExcelSimpleWriter writeTitle(BigTitle bigTitle, String sheetName) {
        if (bigTitle != null) {
            try {
                super.createSheet(sheetName);
                super.writeExecutor.writeTitle(bigTitle);
            } catch (Exception e) {
                super.close();
                throw e;
            }
        }
        return this;
    }

    /**
     * To write
     *
     * @param data Sequential padding, which needs to correspond to the header sequence
     * @return this
     */
    public ExcelSimpleWriter write(List<?> data) {
        return this.write(data, super.defaultSheetName, true);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @return this
     */
    public ExcelSimpleWriter write(List<?> data, String sheetName) {
        return this.write(data, sheetName, true);
    }

    /**
     * To write
     *
     * @param data     data
     * @param needHead need to write the header
     * @return this
     */
    public ExcelSimpleWriter write(List<?> data, boolean needHead) {
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
    public ExcelSimpleWriter write(List<?> data, String sheetName, boolean needHead) {
        try {
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
        } catch (Exception e) {
            super.close();
            throw e;
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
    public ExcelSimpleWriter fieldSelector(BiFunction<Integer, List<Field>, Field> fieldSelector) {
        this.fieldSelector = fieldSelector;
        return this;
    }

    /**
     * Add write listener
     *
     * @param listener Write listener
     * @return this
     */
    public ExcelSimpleWriter listener(ExcelWriteListener listener) {
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
    public ExcelSimpleWriter listener(List<? extends ExcelWriteListener> listeners) {
        if (listeners != null) {
            listeners.forEach(this::listener);
        }
        return this;
    }

    /**
     * Bind the exported Excel file to the currently set unique key,
     * Can be used to {@link ExcelClassReader#check} for a match with an entity class when a file is imported.
     *
     * @param idCard Each exported file recommends that the key be set to be unique.
     *               If empty, the binding is invalid
     * @return this
     */
    public ExcelSimpleWriter bind(String idCard) {
        if (!StringUtils.hasText(idCard)) {
            super.close();
            throw new ExcelException("idCard cannot be empty");
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
    public ExcelSimpleWriter unbind() {
        super.context.setBind(false);
        return this;
    }

    /**
     * Set excel write mode
     *
     * @param mode WRMode
     * @return this
     */
    public ExcelSimpleWriter mode(WRMode mode) {
        super.context.setWrMode(mode);
        return this;
    }
}
