package cn.gjing.excel.base.listener.read;

import java.util.List;

/**
 * Import the data collection listener, and the Excel executor collects the Excel entities generated for each row.
 * Triggered when all data import is complete
 *
 * @author Gjing
 **/
@FunctionalInterface
public interface ExcelResultReadListener<R> extends ExcelReadListener {
    /**
     * Notify the user to take the data
     *
     * @param result Import all the data generated after success
     */
    void notify(List<R> result);
}
