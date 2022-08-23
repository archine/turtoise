package cn.gjing.excel.base.annotation;

import cn.gjing.excel.base.meta.ExcelType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * The declaration is an Excel entity that maps to Excel files.
 * the Excel processor doesn't do anything to normal classes
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {
    /**
     * Name of the Excel file generated when exporting.
     * If this is not set, the current time is used as the name
     *
     * @return Excel file name
     */
    String value() default "";

    /**
     * Excel file type generated when exporting
     *
     * @return Excel file type
     */
    ExcelType type() default ExcelType.XLS;

    /**
     * Window size, which is flushed to disk when exported if the data that has been
     * written out exceeds the specified size. only for xlsx
     *
     * @return windowSize
     */
    int windowSize() default 500;

    /**
     * Number of rows loaded into memory at import time
     * only for xlsx
     *
     * @return cache rows
     */
    int cacheRow() default 100;

    /**
     * Buffer size to use when reading InputStream to file,
     * only for xlsx
     *
     * @return bufferSize
     */
    int bufferSize() default 2048;

    /**
     * Excel header row height
     *
     * @return headHeight
     */
    short headerHeight() default 450;

    /**
     * Excel body row height
     *
     * @return bodyHeight
     */
    short bodyHeight() default 390;

    /**
     * Set the ID card for the exported file.
     *
     * @return key
     */
    String idCard() default "";
}
