package cn.gjing.excel.base.meta;

/**
 * Excel write executor mode
 *
 * @author Gjing
 **/
public enum ExecMode {
    /**
     * Fixed class export
     */
    W_FIXED_CLASS,
    /**
     * Any class export
     */
    W_ANY_CLASS,
    /**
     * Import and generate corresponding class objects
     */
    R_Class
}
