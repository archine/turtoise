package cn.gjing.excel.base.listener;

import cn.gjing.excel.base.meta.ExecMode;

import java.util.List;

/**
 * Excel listener initializer, the base listener from which you can initialize imports and exports,
 * for all import and export methods
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelListenerInitializer {
    /**
     * Initialize the listener list
     *
     * @param excelEntity    Current Excel entity
     * @param execMode       Current execution mode
     * @param excelListeners excel listeners
     */
    void initListeners(Class<?> excelEntity, ExecMode execMode, List<ExcelListener> excelListeners);
}
