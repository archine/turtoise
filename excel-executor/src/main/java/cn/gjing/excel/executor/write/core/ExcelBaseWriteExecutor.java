package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.context.ExcelWriterContext;

import java.util.List;

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
     * Write excel head
     *
     * @param needHead  Whether to set header
     */
    public abstract void writeHead(boolean needHead);

    /**
     * Write excel body
     *
     * @param data Export data
     */
    public abstract void writeBody(List<?> data);
}
