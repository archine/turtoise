package cn.gjing.excel.base.annotation;

import cn.gjing.excel.base.listener.read.ExcelEmptyReadListener;
import cn.gjing.excel.base.meta.ExcelColor;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Declare a field as the Excel header.
 * The actuator does not process normal fields when exporting imports
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {
    /**
     * Fields map to Excel header names.
     * one or more names, each representing one level,
     * are superimposed downwards.
     * all header levels should be consistent within the same Excel entity
     *
     * @return Excel header names
     */
    String[] value() default "";

    /**
     * Specifies the column where the table header resides.
     * the value can be seen in upper case letters at the top of the table header in Excel files.
     * used to distinguish table headers with the same name when importing
     *
     * @return table header resides
     */
    String title() default "";

    /**
     * Header width
     *
     * @return header width, unit (px)
     */
    int width() default 5120;

    /**
     * The table header is set to a sort number, and the Excel executor processes each table header according to the number.
     * If the current Excel class inherits from the parent class,
     * all Excel header fields of the parent class are appended to the current Excel header fields before sorting.
     * If they have the same sequence number, they are sorted according to the order in which the table header fields appear.
     *
     * @return sort index
     */
    int order() default 0;

    /**
     * Set the format of all cells below the current Excel table header when exporting.
     * some commonly used formats are {
     * <p>
     * ------- @ as text
     * ------- 0 as integer
     * ------- 0.00 is two decimal places
     * ------- yyyy-MM-dd  as 年-月-日
     * }
     *
     * See Excel official cell format for more information
     *
     * @return format
     */
    String format() default "";

    /**
     * Whether the body cell below the table header is mandatory.
     * if true, the {@link ExcelEmptyReadListener} will be triggered if the content of the cell is detected as empty during import
     *
     * @return boolean
     */
    boolean required() default false;

    /**
     * Remove Spaces from content cells that are read as text during import
     *
     * @return boolean
     */
    boolean trim() default false;

    /**
     * Excel header color array that can be set separately for each level of header.
     * if the color array is smaller than the header series, the last color is used
     *
     * @return index
     * @see ExcelColor
     */
    ExcelColor[] color() default ExcelColor.LIME;

    /**
     * Excel header font color array that can be set separately for each level of header.
     * if the color array is smaller than the header series, the last color is used
     *
     * @return index
     * @see ExcelColor
     */
    ExcelColor[] fontColor() default ExcelColor.BLACK;
}
