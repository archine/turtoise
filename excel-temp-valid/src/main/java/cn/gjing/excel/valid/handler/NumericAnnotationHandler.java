package cn.gjing.excel.valid.handler;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.valid.ExcelNumericValid;
import cn.gjing.excel.valid.ValidUtil;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * Numeric verification annotation handler
 *
 * @author Gjing
 **/
public class NumericAnnotationHandler extends ValidAnnotationHandler {
    public NumericAnnotationHandler() {
        super(ExcelNumericValid.class);
    }

    @Override
    public void handle(Annotation validAnnotation, ExcelWriterContext writerContext, Field field, Row row, int colIndex, Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        ExcelNumericValid numericValid = (ExcelNumericValid) validAnnotation;
        int firstRow = row.getRowNum() + 1;
        int lastRow = numericValid.rows() == 0 ? firstRow : numericValid.rows() + firstRow - 1;
        DataValidationHelper helper = writerContext.getSheet().getDataValidationHelper();
        DataValidationConstraint numericConstraint = helper.createNumericConstraint(numericValid.type().getType(),
                numericValid.operator().getType(), numericValid.val(), "".equals(numericValid.val2()) ? null : numericValid.val2());
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, colIndex, colIndex);
        DataValidation dataValidation = helper.createValidation(numericConstraint, regions);
        ValidUtil.setErrorBox(dataValidation, numericValid.error(), numericValid.rank(), numericValid.errTitle(), numericValid.errMsg(),
                numericValid.prompt(), numericValid.pTitle(), numericValid.pMsg());
        writerContext.getSheet().addValidationData(dataValidation);
    }
}
