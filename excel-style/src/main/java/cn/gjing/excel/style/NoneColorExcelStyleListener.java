package cn.gjing.excel.style;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelStyleWriteListener;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * Colorless Excel style listener, the header and body are base colors,
 * the only difference is that the header will be bold
 *
 * @author Gjing
 **/
public final class NoneColorExcelStyleListener implements ExcelStyleWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext writerContext;
    private final Map<Integer, CellStyle> titleStyles;
    private CellStyle cellStyle;
    private final Map<String, CellStyle> bodyStyles;

    public NoneColorExcelStyleListener() {
        this.titleStyles = new HashMap<>(8);
        this.bodyStyles = new HashMap<>(16);
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
            titleStyle.setFillForegroundColor(bigTitle.getColor().index);
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
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
    public void setHeadStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex) {
        if (this.cellStyle == null) {
            this.cellStyle = this.writerContext.getWorkbook().createCellStyle();
            Font font = this.writerContext.getWorkbook().createFont();
            font.setBold(true);
            cellStyle.setFont(font);
            StyleUtils.setAlignment(cellStyle);
        }
        if (dataIndex == 0) {
            this.setColumnWidth(property, colIndex);
            if (this.writerContext.isTemplate()) {
                this.writerContext.getSheet().setDefaultColumnStyle(colIndex, this.createFormatStyle(property));
            }
        }
        cell.setCellStyle(this.cellStyle);
    }

    @Override
    public void setBodyStyle(Row row, Cell cell, ExcelFieldProperty property, int dataIndex, int colIndex) {
        if (dataIndex == 0) {
            this.setColumnWidth(property, colIndex);
        }
        cell.setCellStyle(this.createFormatStyle(property));
    }

    private CellStyle createFormatStyle(ExcelFieldProperty property) {
        CellStyle cellStyle = this.bodyStyles.get(property.getFormat());
        if (cellStyle == null) {
            cellStyle = this.writerContext.getWorkbook().createCellStyle();
            StyleUtils.setAlignment(cellStyle);
            if (!property.getFormat().isEmpty()) {
                cellStyle.setDataFormat(this.writerContext.getWorkbook().createDataFormat().getFormat(property.getFormat()));
            }
            this.bodyStyles.put(property.getFormat(), cellStyle);
        }
        return cellStyle;
    }

    private void setColumnWidth(ExcelFieldProperty property, int colIndex) {
        int defaultColumnWidth = this.writerContext.getSheet().getColumnWidth(colIndex);
        if (property.getWidth() > defaultColumnWidth) {
            this.writerContext.getSheet().setColumnWidth(colIndex, property.getWidth());
        }
    }
}
