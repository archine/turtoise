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
 * Gray color style listener, Excel header use gray color, body use basic color.
 * set column width according to {@link ExcelField#width()}, set when writing out the Excel header,
 * set cell format according to {@link ExcelField#format()}
 * <p>
 * When used with the {@link ExcelSheetWriteListener}, all cells outside the header cell and the header column are locked,
 * user can only operate on the cells below the header column.
 *
 * @author Gjing
 **/
public final class GrayColorExcelStyleListener implements ExcelStyleWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext writerContext;
    private final Map<Integer, CellStyle> titleStyles;
    private CellStyle headStyle;
    private final Map<String, CellStyle> bodyStyles;

    public GrayColorExcelStyleListener() {
        this.titleStyles = new HashMap<>(8);
        this.bodyStyles = new HashMap<>(16);
    }

    public GrayColorExcelStyleListener(int bodyStyleCacheCapacity, int titleStyleCacheCapacity) {
        this.titleStyles = new HashMap<>(titleStyleCacheCapacity);
        this.bodyStyles = new HashMap<>(bodyStyleCacheCapacity);
    }

    @Override
    public void setContext(ExcelWriterContext writerContext) {
        this.writerContext = writerContext;
        this.headStyle = writerContext.getWorkbook().createCellStyle();
        this.headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.headStyle.setFillForegroundColor(ExcelColor.GREY_25_PERCENT.index);
        Font font = writerContext.getWorkbook().createFont();
        this.headStyle.setFont(font);
        StyleUtils.setAlignment(this.headStyle);
        StyleUtils.setBorder(this.headStyle, ExcelColor.GREY_40_PERCENT);
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
            this.writerContext.getSheet().setDefaultColumnStyle(columnIndex, StyleUtils.createCacheStyle(property, this.bodyStyles, this.writerContext));
        }
        cell.setCellStyle(this.headStyle);
    }

    @Override
    public void setBodyStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex) {
        cell.setCellStyle(StyleUtils.createCacheStyle(property, this.bodyStyles, this.writerContext));
    }
}
