package cn.gjing.excel.style;

import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelRowWriteListener;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel header merge listener.
 * when adjacent cells have the same content, they are automatically merged.
 *
 * @author Gjing
 **/
public class ExcelHeaderMergeListener implements ExcelRowWriteListener, ExcelWriteContextAware {
    private ExcelWriterContext writerContext;

    public ExcelHeaderMergeListener() {
    }

    @Override
    public void setContext(ExcelWriterContext writerContext) {
        this.writerContext = writerContext;
    }

    @Override
    public void completeRow(Sheet sheet, Row row, Object obj, int index, RowType rowType) {
        if (rowType == RowType.HEAD) {
            if (index + 1 == this.writerContext.getHeaderSeries()) {
                int nextRowIndex;
                int nextColIndex = 1;
                int firstCol = 0;
                int lastCol = 0;
                int firstRow = row.getRowNum() - this.writerContext.getHeaderSeries() + 1;
                int lastRow = firstRow;
                for (int y = firstRow, last = this.writerContext.getHeaderSeries() + firstRow; y < last; y++) {
                    Row currentRow = sheet.getRow(y);
                    nextRowIndex = y + 1;
                    for (int x = 0, colNums = currentRow.getLastCellNum(); x < colNums; x++) {
                        Cell cell = currentRow.getCell(x);
                        if (cell == null) {
                            firstCol = x + 1;
                            lastCol = firstCol;
                            nextColIndex++;
                            continue;
                        }
                        while (colNums > nextColIndex) {
                            Cell cell2 = currentRow.getCell(nextColIndex);
                            if (cell2 == null) {
                                x = nextColIndex;
                                nextColIndex += 2;
                                break;
                            }
                            if (cell.getStringCellValue().equals(cell2.getStringCellValue())) {
                                lastCol = nextColIndex++;
                                continue;
                            }
                            nextColIndex++;
                            break;
                        }
                        while (last > nextRowIndex) {
                            Row nextRow = sheet.getRow(nextRowIndex);
                            Cell nextCell = nextRow.getCell(x);
                            if (nextCell == null || !cell.getStringCellValue().equals(nextCell.getStringCellValue())) {
                                break;
                            }
                            lastRow++;
                            nextRowIndex++;
                        }
                        if (firstCol != lastCol || firstRow != lastRow) {
                            ExcelUtils.merge(this.writerContext.getSheet(), firstCol, lastCol, firstRow, lastRow);
                            x = lastCol;
                            firstCol = x + 1;
                            lastCol = firstCol;
                            lastRow = firstRow;
                            nextRowIndex = y + 1;
                        } else {
                            firstCol = x + 1;
                            lastCol = firstCol;
                        }
                    }
                    firstCol = 0;
                    lastCol = 0;
                    firstRow++;
                    lastRow = firstRow;
                    nextColIndex = 1;
                }
            }
        }
    }
}
