package cn.gjing.excel.base.meta;

import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.listener.ExcelInitializer;
import cn.gjing.excel.base.listener.ExcelListener;

import java.util.List;

/**
 * Global initializer meta
 *
 * @author Gjing
 **/
public enum ExcelInitializerMeta {
    INSTANT;

    ExcelInitializerMeta() {

    }

    private ExcelInitializer excelInitializer;

    /**
     * Set an initializer
     *
     * @param initializer Excel initializer
     */
    public void cache(ExcelInitializer initializer) {
        this.excelInitializer = initializer;
    }

    /**
     * The listener in the initializer is added to the context listener cache.
     * called before each import or export
     *
     * @param excelEntity    Current Excel entity
     * @param execMode       Current execution mode
     * @param excelListeners excel listeners
     */
    public void init(Class<?> excelEntity, ExecMode execMode, List<ExcelListener> excelListeners) {
        if (this.excelInitializer == null) {
            return;
        }
        this.excelInitializer.initListeners(excelEntity, execMode, excelListeners);
    }

    /**
     * Get the global Excel file version
     *
     * @param excelEntity Current Excel entity
     * @param execMode    Current execution mode
     * @return Returning NULL will be set according to {@link Excel#type()}
     */
    public ExcelType initType(Class<?> excelEntity, ExecMode execMode) {
        if (this.excelInitializer == null) {
            return null;
        }
        return this.excelInitializer.initExcelType(excelEntity, execMode);
    }
}
