package cn.gjing.excel.valid;

/**
 * Check the type
 *
 * @author Gjing
 **/
public enum ValidType {
    /**
     * Check the type
     */
    INTEGER(1),
    DECIMAL(2),
    TEXT_LENGTH(6);

    private final int type;

    ValidType(int type) {
        this.type = type;
    }

    public int getType() {
        return type;
    }
}
