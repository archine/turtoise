package cn.gjing.excel.executor.util;

import cn.gjing.excel.base.meta.ExcelType;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author Gjing
 **/
public final class ExcelUtils {

    /**
     * Set cell value
     *
     * @param cell  Current cell
     * @param value Attribute values
     */
    public static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            return;
        }
        if (value instanceof String) {
            cell.setCellValue(value.toString());
            return;
        }
        if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
            return;
        }
        if (value instanceof Enum) {
            cell.setCellValue(value.toString());
            return;
        }
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
            return;
        }
        if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
            return;
        }
        if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
            return;
        }
        throw new IllegalArgumentException("Unsupported data type, you can use a data converter " + value);
    }

    /**
     * Check the file type is excel
     *
     * @param fileName Excel file name
     * @return Return NULL to indicate that it is not an Excel file
     */
    public static ExcelType getExcelType(String fileName) {
        if (fileName == null) {
            return null;
        }
        int pos = fileName.lastIndexOf(".") + 1;
        String extension = fileName.substring(pos);
        if ("xls".equals(extension)) {
            return ExcelType.XLS;
        }
        if ("xlsx".equals(extension)) {
            return ExcelType.XLSX;
        }
        return null;
    }

    /**
     * Merge cells
     *
     * @param sheet    Current sheet
     * @param firstCol First column index
     * @param lastCol  last column index
     * @param firstRow First row index
     * @param LastRow  Last row index
     */
    public static void merge(Sheet sheet, int firstCol, int lastCol, int firstRow, int LastRow) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, LastRow, firstCol, lastCol));
    }

    /**
     * Get cell range address object
     *
     * @param sheet Current sheet
     * @param index address index, start of 0
     * @return CellRangeAddress
     */
    public static CellRangeAddress getCellRangeAddress(Sheet sheet, int index) {
        return sheet.getMergedRegion(index);
    }

    /**
     * Create a sum formula
     *
     * @param firstColIndex Which column start
     * @param firstRowIndex Which row start
     * @param lastColIndex  Which column end
     * @param lastRowIndex  Which row end
     * @return expression
     */
    public static String createSumFormula(int firstColIndex, int firstRowIndex, int lastColIndex, int lastRowIndex) {
        return createFormula("SUM", firstColIndex, firstRowIndex, lastColIndex, lastRowIndex);
    }

    /**
     * Create a formula
     *
     * @param suffix        Formula suffix
     * @param firstColIndex Which column start
     * @param firstRowIndex Which row start
     * @param lastColIndex  Which column end
     * @param lastRowIndex  Which row end
     * @return expression
     */
    public static String createFormula(String suffix, int firstColIndex, int firstRowIndex, int lastColIndex, int lastRowIndex) {
        if (firstRowIndex == lastRowIndex) {
            return suffix + "(" + createFormulaX(firstColIndex, firstRowIndex, lastColIndex) + ")";
        }
        if (firstColIndex == lastColIndex) {
            return suffix + "(" + createFormulaY(firstColIndex, firstRowIndex, lastRowIndex) + ")";
        }
        throw new IllegalArgumentException();
    }

    /**
     * Create a font
     *
     * @param workbook workbook
     * @return Font
     */
    public static Font createFont(Workbook workbook) {
        return workbook.createFont();
    }

    /**
     * Create rich text string
     *
     * @param workbook workbook
     * @param content  Rich text content
     * @return RichTextString
     */
    public static RichTextString createRichText(Workbook workbook, String content) {
        return workbook.getCreationHelper().createRichTextString(content);
    }

    /**
     * Create hyperlink
     *
     * @param workbook workbook
     * @param type     link type
     * @return Hyperlink
     */
    public static Hyperlink createLink(Workbook workbook, HyperlinkType type) {
        return workbook.getCreationHelper().createHyperlink(type);
    }

    /**
     * Determines whether a cell has been merged
     *
     * @param sheet  Current sheet
     * @param row    Current row number
     * @param column current column number
     * @return True is merged
     */
    public static boolean isMerge(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Horizontal creation formula
     *
     * @param startColIndex Start column index
     * @param rowIndex      Row index
     * @param endColIndex   End column index
     * @return If the input parameter is 1,1,10, return $B$1:$K$1
     */
    public static String createFormulaX(int startColIndex, int rowIndex, int endColIndex) {
        char start = (char) ('A' + startColIndex);
        if (endColIndex <= 25) {
            if (endColIndex == 0) {
                return "$" + start + "$" + rowIndex;
            } else {
                char end = (char) (start + endColIndex - 1);
                return "$" + start + "$" + rowIndex + ":$" + end + "$" + rowIndex;
            }
        } else {
            char endPrefix = 'A';
            char endSuffix;
            if ((endColIndex - 25) % 26 == 0) {
                endSuffix = (char) ('A' + 25);
                endPrefix = (char) (endPrefix + (endColIndex - 25) / 26 - 1);
            } else {
                endSuffix = (char) ('A' + (endColIndex - 25) % 26 - 1);
                endPrefix = (char) (endPrefix + (endColIndex - 25) / 26);
            }
            return "$" + start + "$" + rowIndex + ":$" + endPrefix + endSuffix + "$" + rowIndex;
        }
    }

    /**
     * Vertical creation formula
     *
     * @param colIndex colIndex
     * @param startRow start row index
     * @param endRow   end row index
     * @return If the input parameter is 1,2,5, return $B$2:$B$5
     */
    public static String createFormulaY(int colIndex, int startRow, int endRow) {
        if (colIndex < 26) {
            return "$" + (char) ('A' + colIndex) + "$" + startRow + ":$" + (char) ('A' + colIndex) + "$" + endRow;
        }
        String all = ParamUtils.numberToEn(colIndex);
        return "$" + all + "$" + startRow + ":$" + all + "$" + endRow;
    }
}
