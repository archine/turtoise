package cn.gjing.excel.base.context;

import cn.gjing.excel.base.listener.ExcelListener;
import cn.gjing.excel.base.meta.WRMode;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Context objects that exist when Excel imports and exports
 *
 * @author Gjing
 **/
public abstract class AbstractExcelContext {
    /**
     * Current workbook
     */
    @Getter
    @Setter
    protected Workbook workbook;

    /**
     * Current sheet
     */
    @Getter
    @Setter
    protected Sheet sheet;

    /**
     * Import and Export Mode
     */
    @Getter
    @Setter
    protected WRMode wrMode;

    /**
     * Listener cache
     */
    @Getter
    protected final List<ExcelListener> listenerCache;

    protected AbstractExcelContext() {
        this.listenerCache = new ArrayList<>();
    }

    /**
     * Add listener instance to cache
     *
     * @param listener Excel listener
     */
    public void addListener(ExcelListener listener) {
        if (listener != null) {
            this.listenerCache.add(listener);
        }
    }
}