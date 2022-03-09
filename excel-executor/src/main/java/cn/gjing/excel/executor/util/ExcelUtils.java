package cn.gjing.excel.executor.util;

import cn.gjing.excel.base.meta.ExcelType;
import org.apache.poi.ss.usermodel.Cell;

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
}
