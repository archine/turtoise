package cn.gjing.excel.base;

import cn.gjing.excel.base.meta.ExcelColor;
import lombok.*;

import java.lang.reflect.Field;

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
     * Excel headers field
     */
    private Field field;

    /**
     * Array of Excel header names.
     */
    @Builder.Default
    private String[] value = new String[0];

    /**
     * Excel header column
     */
    @Builder.Default
    private String title = "";

    /**
     * Column width of the Excel header
     */
    @Builder.Default
    private int width = 5120;

    /**
     * Header serial number
     */
    @Builder.Default
    private int order = 0;

    /**
     * Cell format
     */
    @Builder.Default
    private String format = "";

    /**
     * Color index array
     */
    @Builder.Default
    private ExcelColor[] color = new ExcelColor[]{ExcelColor.LIME};

    /**
     * Font color index array
     */
    @Builder.Default
    private ExcelColor[] fontColor = new ExcelColor[]{ExcelColor.WHITE};
}
