package cn.gjing.excel.base;

import cn.gjing.excel.base.meta.ExcelColor;
import lombok.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel big title, used to insert a blank line that can fill the content you want
 *
 * @author Gjing
 **/
@Getter
@Setter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public final class BigTitle {
    /**
     * The starting row index, -1 is the latest row added to the current Excel file
     */
    @Builder.Default
    private int firstRow = -1;

    /**
     * How many rows to merge down, less than 1 is automatically set to 1
     */
    @Builder.Default
    private int rowNum = 2;

    /**
     * First col index,the default is to start with the first cell
     */
    @Builder.Default
    private int firstCol = 0;

    /**
     * Last col index, -1 is the number of following excel header size
     */
    @Builder.Default
    private int lastCol = -1;

    /**
     * Style index, if the style of the index exists, it will take the existing one,
     * otherwise it will create a new one
     */
    private int styleIndex;

    /**
     * Fill content
     */
    @Builder.Default
    private Object content = "";

    /**
     * Background color
     */
    @Builder.Default
    private ExcelColor color = ExcelColor.NONE;

    /**
     * The height of each row before merging
     */
    @Builder.Default
    private short rowHeight = 350;

    /**
     * Fill in the color of the content
     */
    @Builder.Default
    private ExcelColor fontColor = ExcelColor.BLACK;

    /**
     * The font height to fill the content
     */
    @Builder.Default
    private short fontHeight = 250;

    /**
     * The content location
     */
    @Builder.Default
    private HorizontalAlignment alignment = HorizontalAlignment.LEFT;

    /**
     * Whether to fill the content in bold
     */
    private boolean bold;

    public static BigTitle of(Object content) {
        return BigTitle.builder()
                .content(content)
                .build();
    }

    public static BigTitle of(Object content, int rowNum) {
        return BigTitle.builder()
                .content(content)
                .rowNum(rowNum)
                .build();
    }

    public static BigTitle of(Object content, int rowNum, int firstRow, int firstCol) {
        return BigTitle.builder()
                .content(content)
                .rowNum(rowNum)
                .firstRow(firstRow)
                .firstCol(firstCol)
                .build();
    }
}
