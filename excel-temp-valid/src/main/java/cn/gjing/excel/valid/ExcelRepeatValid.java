package cn.gjing.excel.valid;

import cn.gjing.excel.base.annotation.ExcelField;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Set the content repetition check macro for Excel header column
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelRepeatValid {
    /**
     * Sets the number of rows of cells below the current Excel header
     *
     * @return rows
     */
    int rows() default 100;

    /**
     * If the number of long text exceeds 15 digits,
     * the Excel file will automatically convert the number after 15 digits to 0,
     * indirectly causing duplicate content check error.
     * Once set to true, the cell of the current column needs to be formatted as text [@] {@link ExcelField#format()}
     *
     * @return boolean
     */
    boolean longNumber() default false;

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
    String errMsg() default "不允许输入重复的内容";

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
