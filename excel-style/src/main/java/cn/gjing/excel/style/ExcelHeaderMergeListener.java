package cn.gjing.excel.style;

import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.listener.write.ExcelRowWriteListener;
import cn.gjing.excel.base.meta.RowType;
import cn.gjing.excel.base.util.ExcelUtils;
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
        if (rowType != RowType.HEAD) {
            return;
        }
        if (index == 0) {
            int startCol = 0;
            int endCol = 0;
            int startRow = row.getRowNum();
            int endRow = row.getRowNum();
            // The depth of the header
            int level;
            // The next header column is indexed
            int nextHeadIndex = 1;
            for (int i = 0; i < this.writerContext.getHeaderSeries(); i++) {
                level = i;
                for (int j = 0, len = this.writerContext.getFieldProperties().size(); j < len; j++) {
                    if (nextHeadIndex - j < 1) {
                        nextHeadIndex++;
                    }
                    if (ExcelUtils.isMerge(sheet, startRow, j)) {
                        startCol = nextHeadIndex;
                        endCol = nextHeadIndex;
                        continue;
                    }
                    // Wide search
                    while (len > nextHeadIndex && this.writerContext.getFieldProperties().get(nextHeadIndex).getValue()[level].equals(this.writerContext.getFieldProperties().get(nextHeadIndex - 1).getValue()[level])) {
                        endCol = nextHeadIndex;
                        nextHeadIndex++;
                        j = endCol;
                    }
                    // Deep search
                    while (this.writerContext.getHeaderSeries() - 1 > level && this.writerContext.getFieldProperties().get(j).getValue()[level].equals(this.writerContext.getFieldProperties().get(j).getValue()[level + 1])) {
                        endRow++;
                        level++;
                    }
                    if (startCol != endCol || startRow != endRow) {
                        ExcelUtils.merge(this.writerContext.getSheet(), startCol, endCol, startRow, endRow);
                        level = i;
                        startCol = nextHeadIndex;
                        endCol = nextHeadIndex;
                        endRow = startRow;
                        continue;
                    }
                    startCol = nextHeadIndex;
                    endCol = nextHeadIndex;
                }
                // Init param
                startCol = 0;
                endCol = 0;
                startRow = startRow + 1;
                endRow = startRow;
                nextHeadIndex = 1;
            }
        }
    }
}
