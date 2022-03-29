package cn.gjing.excel.style.util;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.meta.ExcelColor;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

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

    /**
     * Set the border
     *
     * @param cellStyle cellStyle
     * @param color     border color
     */
    public static void setBorder(CellStyle cellStyle, ExcelColor color) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(color.index);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(color.index);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(color.index);
    }

    /**
     * Create cell styles and cache them
     *
     * @param property ExcelFieldProperty
     * @param styleMap styleMap
     * @param context  ExcelWriterContext
     * @return The style currently created
     */
    public static CellStyle createCacheStyle(ExcelFieldProperty property, Map<String, CellStyle> styleMap, ExcelWriterContext context) {
        CellStyle cellStyle = styleMap.get(property.getFormat());
        if (cellStyle == null) {
            cellStyle = context.getWorkbook().createCellStyle();
            StyleUtils.setAlignment(cellStyle);
            if (!property.getFormat().isEmpty()) {
                cellStyle.setDataFormat(context.getWorkbook().createDataFormat().getFormat(property.getFormat()));
            }
            cellStyle.setLocked(false);
            styleMap.put(property.getFormat(), cellStyle);
        }
        return cellStyle;
    }

    /**
     * Set column width
     *
     * @param property ExcelFieldProperty
     * @param colIndex Column index
     * @param context  ExcelWriterContext
     */
    public static void setColumnWidth(ExcelFieldProperty property, int colIndex, ExcelWriterContext context) {
        int defaultColumnWidth = context.getSheet().getColumnWidth(colIndex);
        if (property.getWidth() > defaultColumnWidth) {
            context.getSheet().setColumnWidth(colIndex, property.getWidth());
        }
    }
}
