package cn.gjing.excel.style;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelSheetWriteListener;
import cn.gjing.excel.base.listener.write.ExcelStyleWriteListener;
import cn.gjing.excel.base.meta.ExcelColor;
import cn.gjing.excel.style.util.StyleUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * Adapt the style listener, header color is affected by {@link ExcelField#color()} {@link ExcelField#fontColor()} color configuration,
 * set column width according to {@link ExcelField#width()}, set when writing out the Excel header,
 * set cell format according to {@link ExcelField#format()}.
 * <p>
 * When used with the {@link ExcelSheetWriteListener}, all cells outside the header cell and the header column are locked,
 * user can only operate on the cells below the header column.
 *
 * @author Gjing
 **/
public final class AdaptExcelStyleListener implements ExcelStyleWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext writerContext;
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

    public AdaptExcelStyleListener(int headStyleCacheCapacity, int bodyStyleCacheCapacity, int titleStyleCacheCapacity) {
        this.titleStyles = new HashMap<>(titleStyleCacheCapacity);
        this.bodyStyles = new HashMap<>(bodyStyleCacheCapacity);
        this.headStyles = new HashMap<>(headStyleCacheCapacity);
    }

    @Override
    public void setContext(ExcelWriterContext writerContext) {
        this.writerContext = writerContext;
    }

    @Override
    public void setTitleStyle(BigTitle bigTitle, Cell cell) {
        CellStyle titleStyle = titleStyles.get(bigTitle.getStyleIndex());
        if (titleStyle == null) {
            titleStyle = this.writerContext.getWorkbook().createCellStyle();
            if (bigTitle.getColor() != ExcelColor.NONE) {
                titleStyle.setFillForegroundColor(bigTitle.getColor().index);
                titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            titleStyle.setAlignment(bigTitle.getAlignment());
            titleStyle.setWrapText(true);
            Font font = this.writerContext.getWorkbook().createFont();
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
    public void setHeadStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex) {
        if (dataIndex == 0) {
            int columnIndex = cell.getColumnIndex();
            StyleUtils.setColumnWidth(property, columnIndex, this.writerContext);
            if (this.writerContext.isTemplate()) {
                this.writerContext.getSheet().setDefaultColumnStyle(columnIndex, StyleUtils.createCacheStyle(property, this.bodyStyles, this.writerContext));
            }
        }
        int colorLen = property.getColor().length;
        int fontColorLen = property.getFontColor().length;
        ExcelColor backgroundColor = property.getColor()[dataIndex < colorLen ? dataIndex : colorLen - 1];
        ExcelColor fontColor = property.getFontColor()[dataIndex < fontColorLen ? dataIndex : fontColorLen - 1];
        String key = backgroundColor.index + "|" + fontColor.index;
        CellStyle cellStyle = this.headStyles.get(key);
        if (cellStyle == null) {
            cellStyle = this.writerContext.getWorkbook().createCellStyle();
            cellStyle.setFillForegroundColor(backgroundColor.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = this.writerContext.getWorkbook().createFont();
            font.setBold(true);
            font.setColor(fontColor.index);
            cellStyle.setFont(font);
            StyleUtils.setBorder(cellStyle, ExcelColor.GREY_40_PERCENT);
            StyleUtils.setAlignment(cellStyle);
            this.headStyles.put(key, cellStyle);
        }
        cell.setCellStyle(cellStyle);
    }

    @Override
    public void setBodyStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex) {
        cell.setCellStyle(StyleUtils.createCacheStyle(property, this.bodyStyles, this.writerContext));
    }
}
