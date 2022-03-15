package cn.gjing.excel.valid;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Set custom macros to Excel header columns
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCustomMacro {
    /**
     * Custom macro formula
     *
     * @return formula
     */
    String formula();

    /**
     * Sets the number of rows of cells below the current Excel header
     *
     * @return rows
     */
    int rows() default 100;

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
