package cn.gjing.excel.base.listener;

import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;

import java.util.List;

/**
 * Excel listener initializer, the base listener from which you can initialize imports and exports,
 * for all import and export methods
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelInitializer {
    /**
     * Initialize the listener list
     *
     * @param excelEntity    Current Excel entity, Null in simple mode
     * @param execMode       Current execution mode
     * @param excelListeners excel listeners
     */
    void initListeners(Class<?> excelEntity, ExecMode execMode, List<ExcelListener> excelListeners);

    /**
     * Initialize the global Excel file version
     *
     * @param excelEntity Current Excel entity, Null in simple mode
     * @param execMode    Current execution mode
     * @return Returning NULL will be set according to {@link Excel#type()}, in simple mode, follow the Settings in Excel factory
     */
    default ExcelType initExcelType(Class<?> excelEntity, ExecMode execMode) {
        return null;
    }
}
