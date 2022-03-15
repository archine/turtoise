package cn.gjing.excel.valid.handler;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.valid.ExcelDateValid;
import cn.gjing.excel.valid.ValidUtil;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * Time verification annotation handler
 *
 * @author Gjing
 **/
public class DateAnnotationHandler extends ValidAnnotationHandler {
    public DateAnnotationHandler() {
        super(ExcelDateValid.class);
    }

    @Override
    public void handle(Annotation validAnnotation, ExcelWriterContext writerContext, Field field, Row row, int colIndex, Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        ExcelDateValid excelDateValid = (ExcelDateValid) validAnnotation;
        DataValidationHelper helper = writerContext.getSheet().getDataValidationHelper();
        DataValidationConstraint dvConstraint;
        if (writerContext.getSheet() instanceof SXSSFSheet) {
            dvConstraint = helper.createDateConstraint(excelDateValid.operator().getType(), "date(" + excelDateValid.val().replaceAll("-", ",") + ")",
                    "date(" + excelDateValid.val2().replaceAll("-", ",") + ")", excelDateValid.format());
        } else {
            dvConstraint = helper.createDateConstraint(excelDateValid.operator().getType(), excelDateValid.val(), excelDateValid.val2(), excelDateValid.format());
        }
        int firstRow = row.getRowNum() + 1;
        int lastRow = excelDateValid.rows() == 0 ? firstRow : excelDateValid.rows() + firstRow - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, colIndex, colIndex);
        DataValidation dataValidation = helper.createValidation(dvConstraint, regions);
        ValidUtil.setErrorBox(dataValidation, excelDateValid.error(), excelDateValid.rank(), excelDateValid.errTitle(), excelDateValid.errMsg(),
                excelDateValid.prompt(), excelDateValid.pTitle(), excelDateValid.pMsg());
        writerContext.getSheet().addValidationData(dataValidation);
    }
}
