package cn.gjing.excel.base;

import cn.gjing.excel.base.meta.ExcelColor;
import lombok.*;

/**
 * Excel filed property
 *
 * @author Gjing
 **/
@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelFieldProperty {
    /**
     * Array of Excel header names.
     * If multiple headers appear in the array, and you want to merge the same header when exporting,
     * you need to activate the multi-header mode before the export is executed.
     * The last one in the array belongs to the real header,
     * and if you are importing a multilevel header Excel file, you need to specify the real header start subscript.
     * The header array size must be the same
     */
    @Builder.Default
    private String[] value = new String[0];

    /**
     * The header above the Excel column that will be concatenated with the real header name ,
     * when imported to match the header in the Excel file to an Excel column header matches the Excel column
     * with the header based on the actual header name
     */
    @Builder.Default
    private String title = "";

    /**
     * Column width of the Excel header
     */
    @Builder.Default
    private int width = 5120;

    /**
     * List header sort defaults to the order in which entity fields appear.
     * If the current class inherits from the parent class,
     * the fields of the parent class are appended to the fields of the current subclass.
     * If the ORDER parameter is set, the order is sorted from smallest to largest
     */
    @Builder.Default
    private int order = 0;

    /**
     * Cell format
     * The cell format of the current column is set when exporting,
     * which defaults to general，
     * Some commonly used formats are {
     * <p>
     * ------- @ as text
     * ------- 0 as integer
     * ------- 0.00 is two decimal places
     * ------- yyyy-MM-dd  as 年-月-日
     * }
     * See Excel official cell format for more information
     */
    @Builder.Default
    private String format = "";

    /**
     * Color index array, If the size of the color array is smaller than the number of series in the header,
     * all subsequent headers will use the last one in the color array
     */
    @Builder.Default
    private ExcelColor[] color = new ExcelColor[]{ExcelColor.LIME};

    /**
     * Font color index array, If the size of the color array is smaller than the number of series in the header,
     * all subsequent headers will use the last one in the color array
     */
    @Builder.Default
    private ExcelColor[] fontColor = new ExcelColor[]{ExcelColor.WHITE};
}
