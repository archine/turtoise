package cn.gjing.excel.base.meta;

import cn.gjing.excel.base.annotation.ExcelField;

/**
 * Import and Export Mode
 *
 * @author Gjing
 **/
public enum WRMode {
    /**
     * Insert columns in sequence according to the order in which the header fields appear
     */
    SORT,

    /**
     * Insert columns based on the index of each table header field
     * @see ExcelField#index()
     */
    INDEX,
}
