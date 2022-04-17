package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.read.ExcelBindReader;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel exports in simple mode, not through mapped entities
 *
 * @author Gjing
 **/
public final class ExcelSimpleWriter extends ExcelBaseWriter {

    public ExcelSimpleWriter(ExcelWriterContext context, int windowSize, HttpServletResponse response) {
        super(context, windowSize, response, ExecMode.SIMPLE_WRITE);
    }

    /**
     * Set the Excel single-level header
     * The order attribute of the generated header field is set to the order in which the elements appear in the header array you pass in, starting at 0
     *
     * @param headNames Excel single-level array of header names
     * @return this
     */
    public ExcelSimpleWriter head(String... headNames) {
        if (headNames == null || headNames.length == 0) {
            throw new ExcelException("excel header names cannot be null");
        }
        List<ExcelFieldProperty> properties = new ArrayList<>(headNames.length);
        for (String headName : headNames) {
            properties.add(ExcelFieldProperty.builder()
                    .order(properties.size())
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
    public ExcelSimpleWriter head(List<String[]> headNames) {
        if (headNames == null || headNames.isEmpty()) {
            throw new ExcelException("excel header names cannot be null");
        }
        List<ExcelFieldProperty> properties = new ArrayList<>(headNames.size());
        for (String[] headName : headNames) {
            properties.add(ExcelFieldProperty.builder()
                    .value(headName)
                    .order(properties.size())
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
    public ExcelSimpleWriter head2(List<ExcelFieldProperty> properties) {
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
    public ExcelSimpleWriter write(List<List<Object>> data) {
        return this.write(data, super.defaultSheetName, true);
    }

    /**
     * To write
     *
     * @param data      data
     * @param sheetName sheet name
     * @return this
     */
    public ExcelSimpleWriter write(List<List<Object>> data, String sheetName) {
        return this.write(data, sheetName, true);
    }

    /**
     * To write
     *
     * @param data     data
     * @param needHead need to write the header
     * @return this
     */
    public ExcelSimpleWriter write(List<List<Object>> data, boolean needHead) {
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
    public ExcelSimpleWriter write(List<List<Object>> data, String sheetName, boolean needHead) {
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
     * Can be used to {@link ExcelBindReader#check} for a match with an entity class when a file is imported.
     *
     * @param key Unique key ,Each exported file recommends that the key be set to be unique.
     *            If empty, the binding is invalid
     * @return this
     */
    public ExcelSimpleWriter bind(String key) {
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
    public ExcelSimpleWriter unbind() {
        super.context.setBind(false);
        return this;
    }

    /**
     * Set the current write position
     *
     * @param startCol col index, based on 0
     * @return this
     */
    public ExcelSimpleWriter withPosition(int startCol) {
        if (startCol < 0) {
            throw new ExcelException("write a column index that cannot be less than 0");
        }
        super.writeExecutor.setPosition(startCol);
        return this;
    }
}
