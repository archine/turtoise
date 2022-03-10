package cn.gjing.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel style utils
 *
 * @author Gjing
 **/
public class StyleUtils {
    /**
     * Sets the cell style coordinates
     *
     * @param cellStyle cellStyle
     */
    public static void setAlignment(CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
    }


}
