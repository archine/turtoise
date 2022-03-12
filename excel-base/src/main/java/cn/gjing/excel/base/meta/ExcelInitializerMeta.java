package cn.gjing.excel.base.meta;

import cn.gjing.excel.base.listener.ExcelListener;
import cn.gjing.excel.base.listener.ExcelListenerInitializer;

import java.util.ArrayList;
import java.util.List;

/**
 * Global initializer meta
 *
 * @author Gjing
 **/
public enum ExcelInitializerMeta {
    INSTANT;

    ExcelInitializerMeta() {
        this.initializers = new ArrayList<>(5);
    }

    private final List<ExcelListenerInitializer> initializers;

    /**
     * Add an initializer
     *
     * @param listenerInitializer Excel listener initializer
     * @return this
     */
    public ExcelInitializerMeta cache(ExcelListenerInitializer listenerInitializer) {
        this.initializers.add(listenerInitializer);
        return this;
    }

    /**
     * The listener in the initializer is added to the context listener cache.
     * called before each import or export
     *
     * @param excelEntity    Current Excel entity
     * @param execMode       Current execution mode
     * @param excelListeners excel listeners
     */
    public void init(Class<?> excelEntity, ExecMode execMode,List<ExcelListener> excelListeners) {
        for (ExcelListenerInitializer initializer : initializers) {
            initializer.initListeners(excelEntity,execMode, excelListeners);
        }
    }
}
