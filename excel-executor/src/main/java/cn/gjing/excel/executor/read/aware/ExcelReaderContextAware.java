package cn.gjing.excel.executor.read.aware;

import cn.gjing.excel.base.aware.ExcelAware;
import cn.gjing.excel.base.context.ExcelReaderContext;

/**
 * ExcelReaderContext loader, through which you can obtain the ExcelReaderContext
 *
 * @author Gjing
 **/
public interface ExcelReaderContextAware<R> extends ExcelAware {
    /**
     * Set excel reader context
     *
     * @param readerContext Excel reader context
     */
    void setContext(ExcelReaderContext<R> readerContext);
}
