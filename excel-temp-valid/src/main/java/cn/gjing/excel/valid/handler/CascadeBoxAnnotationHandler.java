package cn.gjing.excel.valid.handler;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.valid.ExcelCascadeBox;
import cn.gjing.excel.valid.ValidUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * Cascade box valid handler
 *
 * @author Gjing
 **/
public class CascadeBoxAnnotationHandler extends ValidAnnotationHandler {

    public CascadeBoxAnnotationHandler() {
        super(ExcelCascadeBox.class);
    }

    @Override
    public void handle(Annotation validAnnotation, ExcelWriterContext writerContext, Field field, Row row, int colIndex, Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        if (writerContext.getWorkbook().getSheet("subsetSheet") == null) {
            if (cascadeValues == null || cascadeValues.isEmpty()) {
                throw new ExcelException("The cascading drop-down box options cannot be left blank");
            }
            Sheet explicitSheet = writerContext.getWorkbook().createSheet("subsetSheet");
            writerContext.getWorkbook().setSheetHidden(writerContext.getWorkbook().getSheetIndex("subsetSheet"), true);
            for (Map.Entry<String, String[]> valueMap : cascadeValues.entrySet()) {
                Name name = writerContext.getWorkbook().getName(valueMap.getKey());
                if (name == null) {
                    int rowIndex = explicitSheet.getPhysicalNumberOfRows();
                    Row subsetSheetRow = explicitSheet.createRow(rowIndex);
                    subsetSheetRow.createCell(0).setCellValue(valueMap.getKey());
                    for (int i = 0, length = valueMap.getValue().length; i < length; i++) {
                        subsetSheetRow.createCell(i + 1).setCellValue(valueMap.getValue()[i]);
                    }
                    String formula = ExcelUtils.createFormulaX(1, rowIndex + 1, valueMap.getValue().length);
                    name = writerContext.getWorkbook().createName();
                    name.setNameName(valueMap.getKey());
                    name.setRefersToFormula("subsetSheet!" + formula);
                }
            }
        }
        ExcelCascadeBox cascadeBox = (ExcelCascadeBox) validAnnotation;
        char parentIndex = (char) ('A' + Integer.parseInt(cascadeBox.link()));
        DataValidationHelper helper = writerContext.getSheet().getDataValidationHelper();
        DataValidationConstraint constraint;
        CellRangeAddressList regions;
        DataValidation dataValidation;
        int firstRow = row.getRowNum() + 1;
        int lastRow = cascadeBox.rows() == 0 ? firstRow : cascadeBox.rows() + firstRow - 1;
        for (int i = firstRow; i <= lastRow; i++) {
            String forMuaString = "INDIRECT($" + parentIndex + "$" + (i + 1) + ")";
            constraint = helper.createFormulaListConstraint(forMuaString);
            regions = new CellRangeAddressList(i, i, colIndex, colIndex);
            dataValidation = helper.createValidation(constraint, regions);
            ValidUtil.setErrorBox(dataValidation, cascadeBox.error(), cascadeBox.rank(), cascadeBox.errTitle(), cascadeBox.errMsg(), cascadeBox.prompt(), cascadeBox.pTitle(), cascadeBox.pMsg());
            writerContext.getSheet().addValidationData(dataValidation);
        }
    }
}
