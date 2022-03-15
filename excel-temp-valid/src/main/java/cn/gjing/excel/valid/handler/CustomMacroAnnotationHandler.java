package cn.gjing.excel.valid.handler;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.valid.ExcelCustomMacro;
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
 * Custom macro annotation handler
 *
 * @author Gjing
 **/
public class CustomMacroAnnotationHandler extends ValidAnnotationHandler {
    public CustomMacroAnnotationHandler() {
        super(ExcelCustomMacro.class);
    }

    @Override
    public void handle(Annotation validAnnotation, ExcelWriterContext writerContext, Field field, Row row, int colIndex, Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        ExcelCustomMacro customMacro = (ExcelCustomMacro) validAnnotation;
        int firstRow = row.getRowNum() + 1;
        int lastRow = customMacro.rows() == 0 ? firstRow : customMacro.rows() + firstRow - 1;
        DataValidationHelper helper = writerContext.getSheet().getDataValidationHelper();
        DataValidationConstraint customConstraint = helper.createCustomConstraint(customMacro.formula());
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, colIndex, colIndex);
        DataValidation validation = helper.createValidation(customConstraint, regions);
        ValidUtil.setErrorBox(validation, customMacro.error(), customMacro.rank(), customMacro.errTitle(), customMacro.errMsg(),
                customMacro.prompt(), customMacro.pTitle(), customMacro.pMsg());
        writerContext.getSheet().addValidationData(validation);
    }
}
