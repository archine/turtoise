package cn.gjing.excel.executor.util;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelField;
import cn.gjing.excel.base.util.ParamUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Bean tools
 *
 * @author Gjing
 **/
public final class BeanUtils {
    /**
     * Set the value of a field of an object
     *
     * @param o     object
     * @param field field
     * @param value value
     */
    public static void setFieldValue(Object o, Field field, Object value) {
        try {
            field.setAccessible(true);
            field.set(o, value);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     * Gets the value in the field
     *
     * @param o     object
     * @param field field
     * @return Object
     */
    public static Object getFieldValue(Object o, Field field) {
        try {
            field.setAccessible(true);
            return field.get(o);
        } catch (IllegalAccessException e) {
            return null;
        }
    }

    /**
     * Get all excel field properties of the parent and child classes
     *
     * @param excelClass Excel mapped entity
     * @param ignores    The exported field is to be ignored
     * @return Excel filed properties
     */
    public static List<ExcelFieldProperty> getExcelFiledProperties(Class<?> excelClass, String[] ignores) {
        List<ExcelFieldProperty> fieldProperties = new ArrayList<>();
        getAllFields(excelClass).stream()
                .filter(e -> e.isAnnotationPresent(ExcelField.class))
                .forEach(e -> {
                    ExcelField excelField = e.getAnnotation(ExcelField.class);
                    String[] headNameArray = excelField.value();
                    for (String name : headNameArray) {
                        if (ParamUtils.contains(ignores, name)) {
                            return;
                        }
                    }
                    fieldProperties.add(ExcelFieldProperty.builder()
                            .value(excelField.value())
                            .field(e)
                            .width(excelField.width())
                            .index(excelField.index())
                            .format(excelField.format())
                            .color(excelField.color())
                            .fontColor(excelField.fontColor())
                            .build());
                });
        return fieldProperties;
    }

    /**
     * Get all fields of the parent and child classes
     *
     * @param clazz Class
     * @return Field list
     */
    public static List<Field> getAllFields(Class<?> clazz) {
        if (clazz == null) {
            return new ArrayList<>();
        }
        Field[] declaredFields = clazz.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>(Arrays.asList(declaredFields));
        Class<?> superclass = clazz.getSuperclass();
        while (superclass != Object.class) {
            fieldList.addAll(Arrays.asList(superclass.getDeclaredFields()));
            superclass = superclass.getSuperclass();
        }
        return fieldList;
    }
}
