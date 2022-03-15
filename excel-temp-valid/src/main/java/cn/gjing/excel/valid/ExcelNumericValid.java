package cn.gjing.excel.valid;

import java.lang.annotation.*;

/**
 * Set value checking macros for Excel header columns
 *
 * @author Gjing
 **/
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelNumericValid {
    /**
     * Sets the number of rows of cells below the current Excel header
     *
     * @return rows
     */
    int rows() default 100;

    /**
     * Operator type
     *
     * @return OperatorType
     */
    OperatorType operator();

    /**
     * Value type
     *
     * @return TEXT_LENGTH
     */
    ValidType type();

    /**
     * Value to be evaluated by the operation type
     *
     * @return expr1
     */
    String val();

    /**
     * Value, which is calculated by the operation type.
     * This parameter is required only when the operation type {@link OperatorType#BETWEEN} or {@link OperatorType#NOT_BETWEEN}
     *
     * @return expr2
     */
    String val2() default "";

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
    String errMsg() default "填写的内容不符合要求";

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
