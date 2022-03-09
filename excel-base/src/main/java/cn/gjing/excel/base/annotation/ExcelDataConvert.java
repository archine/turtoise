package cn.gjing.excel.base.annotation;

import java.lang.annotation.*;

/**
 * Data conversion annotations that process the contents of cells under a table header during import and export
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelDataConvert {
    /**
     * EL expressions that process data when exporting
     *
     * @return EL expression
     */
    String writeExpr() default "";

    /**
     * EL expressions that process data when importing
     *
     * @return EL expression
     */
    String readExpr() default "";
}
