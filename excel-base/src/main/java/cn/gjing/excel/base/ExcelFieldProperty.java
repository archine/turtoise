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
@ToString
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
     * Column width of the Excel header
     */
    @Builder.Default
    private int width = 5120;

    /**
     * Excel header column index
     **/
    @Builder.Default
    private int index = 0;

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
    private ExcelColor[] fontColor = new ExcelColor[]{ExcelColor.BLACK};

    public static ExcelFieldProperty of(String... value) {
        return ExcelFieldProperty.builder()
                .value(value)
                .build();
    }

    public static ExcelFieldProperty of(String format, int index, String... value) {
        return ExcelFieldProperty.builder()
                .value(value)
                .index(index)
                .format(format)
                .build();
    }

    public static ExcelFieldProperty of(int index, String... value) {
        return ExcelFieldProperty.builder()
                .value(value)
                .index(index)
                .build();
    }

    public static ExcelFieldProperty of(String format, String[] header) {
        return ExcelFieldProperty.builder()
                .format(format)
                .value(header)
                .build();
    }
}
