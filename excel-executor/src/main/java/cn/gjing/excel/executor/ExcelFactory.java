package cn.gjing.excel.executor;

import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.executor.read.ExcelClassReader;
import cn.gjing.excel.executor.util.BeanUtils;
import cn.gjing.excel.executor.write.ExcelFixedClassWriter;
import cn.gjing.excel.executor.write.ExcelAnyClassWriter;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Objects;

/**
 * Create excel reader and writer，Used to import and export Excel
 *
 * @author Gjing
 **/
public final class ExcelFactory {
    private ExcelFactory() {
    }

    /**
     * Create an Excel fixed class writer
     *
     * @param excelEntity Excel entity
     * @param response    response
     * @param ignores     Which table heads to be ignored when exporting, in the case of multiple table heads,
     *                    there are more than one child table heads under the ignored table head,
     *                    then the child table head will be ignored, if the ignored table head is from the table head
     *                    then it is ignored
     * @return ExcelWriter
     */
    public static ExcelFixedClassWriter createWriter(Class<?> excelEntity, HttpServletResponse response, String... ignores) {
        return createWriter(null, excelEntity, response, ignores);
    }

    /**
     * Create an Excel fixed class writer
     *
     * @param fileName    Excel file name，The priority is higher than the annotation specification
     * @param excelEntity Excel entity
     * @param response    response
     * @param ignores     The name of the header to be ignored when exporting.
     *                    If it is a parent, all children below it will be ignored as well
     * @return ExcelWriter
     */
    public static ExcelFixedClassWriter createWriter(String fileName, Class<?> excelEntity, HttpServletResponse response, String... ignores) {
        Objects.requireNonNull(excelEntity, "Excel mapping class cannot be null");
        Excel excel = excelEntity.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "@Excel annotation was not found on the " + excelEntity);
        ExcelWriterContext context = new ExcelWriterContext();
        context.setExcelEntity(excelEntity);
        context.setExcelType(excel.type());
        context.setFieldProperties(BeanUtils.getExcelFiledProperties(excelEntity, ignores));
        context.setFileName(StringUtils.hasText(fileName) ? fileName : "".equals(excel.value()) ? LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) : excel.value());
        context.setHeaderHeight(excel.headerHeight());
        context.setHeaderSeries(context.getFieldProperties().size() == 0 ? 0 : context.getFieldProperties().get(0).getValue().length);
        context.setBodyHeight(excel.bodyHeight());
        context.setUniqueKey("".equals(excel.uniqueKey()) ? excelEntity.getName() : excel.uniqueKey());
        return new ExcelFixedClassWriter(context, excel, response);
    }

    /**
     * Create an Excel any class writer
     *
     * @param fileName Excel file name
     * @param response response
     * @return ExcelSimpleWriter
     */
    public static ExcelAnyClassWriter createAnyClassWriter(String fileName, HttpServletResponse response) {
        return createAnyClassWriter(fileName, response, ExcelType.XLS, 500);
    }

    /**
     * Create an Excel any class writer
     *
     * @param fileName  Excel file name
     * @param response  response
     * @param excelType Excel file type
     * @return ExcelSimpleWriter
     */
    public static ExcelAnyClassWriter createAnyClassWriter(String fileName, HttpServletResponse response, ExcelType excelType) {
        return createAnyClassWriter(fileName, response, excelType, 500);
    }

    /**
     * Create an Excel any writer
     *
     * @param fileName   Excel file name
     * @param response   response
     * @param excelType  Excel file type
     * @param windowSize Window size, which is flushed to disk when exported
     *                   if the data that has been written out exceeds the specified size
     *                   only for xlsx
     * @return ExcelSimpleWriter
     */
    public static ExcelAnyClassWriter createAnyClassWriter(String fileName, HttpServletResponse response, ExcelType excelType, int windowSize) {
        ExcelWriterContext context = new ExcelWriterContext();
        context.setFileName(StringUtils.hasText(fileName) ? fileName : LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
        context.setExcelEntity(null);
        context.setExcelType(excelType);
        context.setBind(false);
        return new ExcelAnyClassWriter(context, windowSize, response);
    }

    /**
     * Create an Excel class reader
     *
     * @param file       Excel file
     * @param excelClass Object class to be generated
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelClassReader<R> createReader(MultipartFile file, Class<R> excelClass) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getOriginalFilename());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createReader(file.getInputStream(), excelClass, excelType);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel class reader
     *
     * @param file       Excel file
     * @param excelClass Object class to be generated
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelClassReader<R> createReader(File file, Class<R> excelClass) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getName());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createReader(Files.newInputStream(file.toPath()), excelClass, excelType);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel class reader
     *
     * @param inputStream Excel file inputStream
     * @param excelClass Object class to be generated
     * @param excelType   Excel file type
     * @param <R>         Entity type
     * @return ExcelReader
     */
    public static <R> ExcelClassReader<R> createReader(InputStream inputStream, Class<R> excelClass, ExcelType excelType) {
        Objects.requireNonNull(excelClass, "Excel mapping class cannot be null");
        Excel excel = excelClass.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "@Excel annotation was not found on the " + excel);
        ExcelReaderContext<R> readerContext = new ExcelReaderContext<>(excelClass);
        readerContext.setUniqueKey("".equals(excel.uniqueKey()) ? excelClass.getName() : excel.uniqueKey());
        readerContext.setFieldProperties(BeanUtils.getExcelFiledProperties(excelClass, null));
        return new ExcelClassReader<>(readerContext, inputStream, excelType, excel.cacheRow(), excel.bufferSize());
    }
}
