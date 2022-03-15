package cn.gjing.excel.valid;

import java.lang.annotation.*;

/**
 * Set time checking macros for Excel header columns
 *
 * @author Gjing
 **/
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelDateValid {
    /**
     * Sets the number of rows of cells below the current Excel header
     *
     * @return rows
     */
    int rows() default 100;

    /**
     * Check time format
     *
     * @return expr
     */
    String format() default "yyyy-MM-dd";

    /**
     * Operator type
     *
     * @return OperatorType
     */
    OperatorType operator() default OperatorType.BETWEEN;

    /**
     * Value to be evaluated by the operation type
     *
     * @return val
     */
    String val() default "1900-01-01";

    /**
     * Value, which is calculated by the operation type.
     * This parameter is required only when the operation type {@link OperatorType#BETWEEN} or {@link OperatorType#NOT_BETWEEN}
     *
     * @return val2
     */
    String val2() default "2999-01-01";

    /**
     * The entered content does not meet the conditions. Open the error box
     *
     * @return True is open
     */
    boolean error() default true;

    /**
     * The style level of the error box
     *
     * @return level
     */
    Rank rank() default Rank.STOP;

    /**
     * The title of the error box
     *
     * @return title
     */
    String errTitle() default "";

    /**
     * The contents of the error box
     *
     * @return content
     */
    String errMsg() default "填写的时间不满足要求";

    /**
     * Enter the content to open the prompt box
     *
     * @return false
     */
    boolean prompt() default false;

    /**
     * Title of the prompt box
     *
     * @return ""
     */
    String pTitle() default "";

    /**
     * Prompt content
     *
     * @return ""
     */
    String pMsg() default "";
}
