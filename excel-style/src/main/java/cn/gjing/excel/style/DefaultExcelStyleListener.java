package cn.gjing.excel.style;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelStyleWriteListener;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * The default style listener, header color is affected by {@link ExcelField#color()} color configuration
 *
 * @author Gjing
 **/
public final class DefaultExcelStyleListener implements ExcelStyleWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext context;
    private final Map<Integer, CellStyle> titleStyles;
    private final Map<Class<?>, Map<Integer, List<CellStyle>>> headStyleData;
    private final Map<String, CellStyle> bodyStyles;

    public DefaultExcelStyleListener() {
        this.headStyleData = new HashMap<>(8);
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
            titleStyle.setFillForegroundColor(bigTitle.getColor().index);
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
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
        Map<Integer, List<CellStyle>> headStyle = this.headStyleData.computeIfAbsent(this.context.getExcelEntity(), k -> new HashMap<>(32));
        List<CellStyle> cellStyleList = headStyle.get(colIndex);
        if (cellStyleList == null) {
            cellStyleList = new ArrayList<>();
            int colorLength = property.getColor().length;
            int fontColorLength = property.getFontColor().length;
            CellStyle cellStyle;
            int maxIndex = Math.max(colorLength, fontColorLength);
            for (int i = 0; i < maxIndex; i++) {
                cellStyle = this.context.getWorkbook().createCellStyle();
                cellStyle.setFillForegroundColor(property.getColor()[colorLength > i ? i : colorLength - 1].index);
                Font font = this.context.getWorkbook().createFont();
                font.setBold(true);
                font.setColor(property.getFontColor()[fontColorLength > i + 1 ? i : fontColorLength - 1].index);
                cellStyle.setFont(font);
                this.setColorAndBorder(cellStyle);
                StyleUtils.setAlignment(cellStyle);
                cellStyleList.add(cellStyle);
            }
            headStyle.put(colIndex, cellStyleList);
        }
        if (dataIndex == 0) {
            this.setColumnWidth(property, colIndex);
            if (this.context.isTemplate()) {
                this.context.getSheet().setDefaultColumnStyle(colIndex, this.createBodyStyle(property));
            }
        }
        cell.setCellStyle(cellStyleList.size() > dataIndex ? cellStyleList.get(dataIndex) : cellStyleList.get(cellStyleList.size() - 1));
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

    private void setColorAndBorder(CellStyle cellStyle) {
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.index);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.index);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.GREY_40_PERCENT.index);
    }
}
