package cn.gjing.excel.style;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelStyleWriteListener;
import cn.gjing.excel.base.meta.ExcelColor;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * Adapt the style listener, header color is affected by {@link ExcelField#color()} {@link ExcelField#fontColor()} color configuration,
 * excel header font bold,
 * set column width according to {@link ExcelField#width()},
 * set cell format according to {@link ExcelField#format()},
 *
 * @author Gjing
 **/
public final class AdaptExcelStyleListener implements ExcelStyleWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext context;
    /**
     * Big title style cache, key for style index
     */
    private final Map<Integer, CellStyle> titleStyles;
    /**
     * Header style cache, key is a string of font and background color combination
     */
    private final Map<String, CellStyle> headStyles;
    /**
     * Body style cache, key in cell format
     */
    private final Map<String, CellStyle> bodyStyles;

    public AdaptExcelStyleListener() {
        this.headStyles = new HashMap<>(16);
        this.bodyStyles = new HashMap<>(16);
        this.titleStyles = new HashMap<>(8);
    }

    @Override
    public void setContext(ExcelWriterContext writerContext) {
        this.context = writerContext;
    }

    @Override
    public void setTitleStyle(BigTitle bigTitle, Cell cell) {
        CellStyle titleStyle = titleStyles.get(bigTitle.getStyleIndex());
        if (titleStyle == null) {
            titleStyle = this.context.getWorkbook().createCellStyle();
            if (bigTitle.getColor() != ExcelColor.NONE) {
                titleStyle.setFillForegroundColor(bigTitle.getColor().index);
                titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            titleStyle.setAlignment(bigTitle.getAlignment());
            titleStyle.setWrapText(true);
            Font font = this.context.getWorkbook().createFont();
            font.setColor(bigTitle.getFontColor().index);
            font.setBold(bigTitle.isBold());
            font.setFontHeight(bigTitle.getFontHeight());
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            titleStyle.setFont(font);
            this.titleStyles.put(bigTitle.getStyleIndex(), titleStyle);
        }
        cell.setCellStyle(titleStyle);
    }

    @Override
    public void setHeadStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex) {
        int colorLen = property.getColor().length;
        int fontColorLen = property.getFontColor().length;
        ExcelColor backgroundColor = property.getColor()[dataIndex < colorLen ? dataIndex : colorLen - 1];
        ExcelColor fontColor = property.getFontColor()[dataIndex < fontColorLen ? dataIndex : fontColorLen - 1];
        String key = backgroundColor.index + "|" + fontColor.index;
        CellStyle cellStyle = this.headStyles.get(key);
        if (cellStyle == null) {
            cellStyle = this.context.getWorkbook().createCellStyle();
            cellStyle.setFillForegroundColor(backgroundColor.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = this.context.getWorkbook().createFont();
            font.setBold(true);
            font.setColor(fontColor.index);
            cellStyle.setFont(font);
            StyleUtils.setBorder(cellStyle, ExcelColor.GREY_40_PERCENT);
            StyleUtils.setAlignment(cellStyle);
            this.headStyles.put(key, cellStyle);
        }
        if (dataIndex == 0) {
            this.setColumnWidth(property, colIndex);
            if (this.context.isTemplate()) {
                this.context.getSheet().setDefaultColumnStyle(colIndex, this.createBodyStyle(property));
            }
        }
        cell.setCellStyle(cellStyle);
    }

    @Override
    public void setBodyStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex) {
        if (dataIndex == 0) {
            this.setColumnWidth(property, colIndex);
        }
        cell.setCellStyle(this.createBodyStyle(property));
    }

    private CellStyle createBodyStyle(ExcelFieldProperty property) {
        CellStyle cellStyle = this.bodyStyles.get(property.getFormat());
        if (cellStyle == null) {
            cellStyle = this.context.getWorkbook().createCellStyle();
            StyleUtils.setAlignment(cellStyle);
            if (!property.getFormat().isEmpty()) {
                cellStyle.setDataFormat(this.context.getWorkbook().createDataFormat().getFormat(property.getFormat()));
            }
            this.bodyStyles.put(property.getFormat(), cellStyle);
        }
        return cellStyle;
    }

    private void setColumnWidth(ExcelFieldProperty property, int colIndex) {
        int defaultColumnWidth = this.context.getSheet().getColumnWidth(colIndex);
        if (property.getWidth() > defaultColumnWidth) {
            this.context.getSheet().setColumnWidth(colIndex, property.getWidth());
        }
    }
}
