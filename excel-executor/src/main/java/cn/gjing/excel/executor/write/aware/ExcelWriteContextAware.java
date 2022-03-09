package cn.gjing.excel.executor.write.aware;

import cn.gjing.excel.base.aware.ExcelAware;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;

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
