package cn.gjing.excel.base.listener;

import java.util.List;

/**
 * Excel listener initializer, the base listener from which you can initialize imports and exports,
 * for all import and export methods
 *
 * @author Gjing
 **/
public interface ExcelListenerInitializer {
    /**
     * Initialize the listener list
     *
     * @param excelListeners excel listeners
     */
    void initListeners(List<ExcelListener> excelListeners);
}
