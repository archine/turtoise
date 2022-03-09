package cn.gjing.excel.base.meta;

/**
 * Excel file type
 *
 * @author Gjing
 **/
public enum ExcelType {
    /**
     * 2003 version, more than 6w data will report error or OOM,
     */
    XLS,

    /**
     * 2007 version,used for big data processing,
     */
    XLSX
}
