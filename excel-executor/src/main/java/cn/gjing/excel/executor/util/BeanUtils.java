package cn.gjing.excel.executor.util;

import cn.gjing.excel.base.ExcelFieldProperty;
import cn.gjing.excel.base.annotation.ExcelField;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

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
     * Get all excel fields of the parent and child classes
     *
     * @param excelClass      Excel mapped entity
     * @param ignores         The exported field is to be ignored
     * @param fieldProperties Excel field property
     * @return Excel fields
     */
    public static List<Field> getExcelFields(Class<?> excelClass, String[] ignores, List<ExcelFieldProperty> fieldProperties) {
        List<Field> fieldList = getAllFields(excelClass);
        fieldList = fieldList.stream()
                .filter(e -> e.isAnnotationPresent(ExcelField.class))
                .sorted(Comparator.comparing(e -> e.getAnnotation(ExcelField.class).order()))
                .filter(e -> {
                    ExcelField excelField = e.getAnnotation(ExcelField.class);
                    String[] headNameArray = excelField.value();
                    for (String name : headNameArray) {
                        if (ParamUtils.contains(ignores, name)) {
                            return false;
                        }
                    }
                    if (fieldProperties != null) {
                        fieldProperties.add(ExcelFieldProperty.builder()
                                .value(excelField.value())
                                .title(excelField.title())
                                .width(excelField.width())
                                .order(excelField.order())
                                .format(excelField.format())
                                .color(excelField.color())
                                .fontColor(excelField.fontColor())
                                .build());
                    }
                    return true;
                })
                .collect(Collectors.toList());
        return fieldList;
    }

    /**
     * Generate excel field Map
     *
     * @param excelClass Excel mapped entity
     * @return Excel field map
     */
    public static Map<String, Field> getExcelFieldsMap(Class<?> excelClass) {
        List<Field> fieldList = getAllFields(excelClass);
        return fieldList.stream()
                .filter(e -> e.isAnnotationPresent(ExcelField.class))
                .collect(Collectors.toMap(e -> {
                    ExcelField excelField = e.getAnnotation(ExcelField.class);
                    String[] headArray = excelField.value();
                    return headArray[headArray.length - 1] + excelField.title();
                }, f -> f));
    }

    /**
     * Get all fields of the parent and child classes
     *
     * @param clazz Class
     * @return Field list
     */
    public static List<Field> getAllFields(Class<?> clazz) {
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
