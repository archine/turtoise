package cn.gjing.excel.valid.handler;

import cn.gjing.excel.base.context.ExcelWriterContext;
import org.apache.poi.ss.usermodel.Row;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Validates annotation processor metadata
 *
 * @author Gjing
 **/
public enum HandleMeta {
    INSTANCE;

    private final List<ValidAnnotationHandler> handlers = new ArrayList<>();

    HandleMeta() {
        this.handlers.add(new DropdownAnnotationHandler());
        this.handlers.add(new NumericAnnotationHandler());
        this.handlers.add(new RepeatAnnotationHandler());
        this.handlers.add(new DateAnnotationHandler());
        this.handlers.add(new CascadeBoxAnnotationHandler());
        this.handlers.add(new CustomMacroAnnotationHandler());
    }

    public void exec(Field field, ExcelWriterContext writerContext, Row row, int colIndex, Map<String, String[]> boxValues, Map<String, String[]> cascadeValues) {
        for (ValidAnnotationHandler handler : handlers) {
            Annotation annotation = field.getAnnotation(handler.getAnnotationClass());
            if (annotation != null) {
                handler.handle(annotation, writerContext, field, row, colIndex, boxValues, cascadeValues);
                break;
            }
        }
    }
}
