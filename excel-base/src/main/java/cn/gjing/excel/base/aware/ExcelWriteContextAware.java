package cn.gjing.excel.base.aware;

import cn.gjing.excel.base.context.ExcelWriterContext;

/**
 * ExcelWriterContext loader, through which you can obtain the ExcelWriterContext
 *
 * @author Gjing
 **/
public interface ExcelWriteContextAware extends ExcelAware {
    /**
     * Set excel writer context
     *
     * @param writerContext Excel writer context
     */
    void setContext(ExcelWriterContext writerContext);
}
