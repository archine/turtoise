package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.executor.write.context.ExcelWriterContext;

import java.util.List;
import java.util.Map;

/**
 * Excel writes the core processor
 *
 * @author Gjing
 **/
public abstract class ExcelBaseWriteExecutor {
    protected final ExcelWriterContext context;

    public ExcelBaseWriteExecutor(ExcelWriterContext context) {
        this.context = context;
    }

    /**
     * Set excel head
     *
     * @param needHead  Whether to set header
     * @param boxValues Excel dropdown box value
     */
    public abstract void writeHead(boolean needHead, Map<String, String[]> boxValues);

    /**
     * Set excel body
     *
     * @param data Export data
     */
    public abstract void writeBody(List<?> data);
}
