package cn.gjing.excel.base.meta;

/**
 * Excel write executor mode
 *
 * @author Gjing
 **/
public enum ExecMode {
    /**
     * Bind export
     */
    W_BIND,
    /**
     * Simple export
     */
    W_SIMPLE,
    /**
     * Import and generate corresponding class objects
     */
    R_CLASS
}
