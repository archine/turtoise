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
        if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
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
        sheet.addMergedRegionUnsafe(new CellRangeAddress(firstRow, LastRow, firstCol, lastCol));
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
     * Horizontal creation formula
     *
     * @param startColIndex Start column index
     * @param rowIndex      Row index
     * @param endColIndex   End column index
     * @return If the input parameter is 1,1,10, return $B$1:$K$1
     */
    public static String createFormulaX(int startColIndex, int rowIndex, int endColIndex) {
        return "$" + ParamUtils.numberToEn(startColIndex) + "$" + rowIndex + ":$" + ParamUtils.numberToEn(endColIndex) + "$" + rowIndex;
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
        String colEn = ParamUtils.numberToEn(colIndex);
        return "$" + colEn + "$" + startRow + ":$" + colEn + "$" + endRow;
    }
}
