package cn.gjing.excel.base.annotation;

import cn.gjing.excel.base.exception.ExcelAssertException;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * When this expression is imported,
 * it puts the contents of all cells below the Excel header field into the expression for calculation.
 * If it is false, thrown an {@link ExcelAssertException}
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAssert {
    /**
     * EL expression. The result must be of type bool
     *
     * @return boolean expression
     */
    String expr();

    /**
     * If the result is false, the exception message is thrown
     *
     * @return message
     */
    String message() default "Invalid content";
}
