package cn.gjing.excel.executor;

import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.meta.ExcelInitializerMeta;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.util.BeanUtils;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.executor.read.ExcelBindReader;
import cn.gjing.excel.executor.read.ExcelSimpleReader;
import cn.gjing.excel.executor.write.ExcelBindWriter;
import cn.gjing.excel.executor.write.ExcelSimpleWriter;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
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
     * Create an Excel bind writer
     *
     * @param excelEntity Excel entity
     * @param response    response
     * @param ignores     Which table heads to be ignored when exporting, in the case of multiple table heads,
     *                    there are more than one child table heads under the ignored table head,
     *                    then the child table head will be ignored, if the ignored table head is from the table head
     *                    then it is ignored
     * @return ExcelWriter
     */
    public static ExcelBindWriter createWriter(Class<?> excelEntity, HttpServletResponse response, String... ignores) {
        return createWriter(null, excelEntity, response, ignores);
    }

    /**
     * Create an Excel writer
     *
     * @param fileName    Excel file name，The priority is higher than the annotation specification
     * @param excelEntity Excel entity
     * @param response    response
     * @param ignores     The name of the header to be ignored when exporting.
     *                    If it is a parent, all children below it will be ignored as well
     * @return ExcelWriter
     */
    public static ExcelBindWriter createWriter(String fileName, Class<?> excelEntity, HttpServletResponse response, String... ignores) {
        Objects.requireNonNull(excelEntity, "Excel mapping class cannot be null");
        Excel excel = excelEntity.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "@Excel annotation was not found on the " + excelEntity);
        ExcelWriterContext context = new ExcelWriterContext();
        context.setExcelEntity(excelEntity);
        context.setFieldProperties(BeanUtils.getExcelFiledProperties(excelEntity, ignores));
        context.setExcelType(ExcelInitializerMeta.INSTANT.initType(excelEntity, ExecMode.WRITE) == null ? excel.type() : ExcelInitializerMeta.INSTANT.initType(excelEntity, ExecMode.WRITE));
        context.setFileName(StringUtils.hasText(fileName) ? fileName : "".equals(excel.value()) ? LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) : excel.value());
        context.setHeaderHeight(excel.headerHeight());
        context.setHeaderSeries(context.getFieldProperties().get(0).getValue().length);
        context.setBodyHeight(excel.bodyHeight());
        context.setUniqueKey("".equals(excel.uniqueKey()) ? excelEntity.getName() : excel.uniqueKey());
        return new ExcelBindWriter(context, excel, response);
    }

    /**
     * Create an Excel simple writer
     *
     * @param fileName Excel file name
     * @param response response
     * @return ExcelSimpleWriter
     */
    public static ExcelSimpleWriter createSimpleWriter(String fileName, HttpServletResponse response) {
        return createSimpleWriter(fileName, response, ExcelType.XLS, 500);
    }

    /**
     * Create an Excel simple writer
     *
     * @param fileName  Excel file name
     * @param response  response
     * @param excelType Excel file type
     * @return ExcelSimpleWriter
     */
    public static ExcelSimpleWriter createSimpleWriter(String fileName, HttpServletResponse response, ExcelType excelType) {
        return createSimpleWriter(fileName, response, excelType, 500);
    }

    /**
     * Create an Excel simple writer
     *
     * @param fileName   Excel file name
     * @param response   response
     * @param excelType  Excel file type
     * @param windowSize Window size, which is flushed to disk when exported
     *                   if the data that has been written out exceeds the specified size
     *                   only for xlsx
     * @return ExcelSimpleWriter
     */
    public static ExcelSimpleWriter createSimpleWriter(String fileName, HttpServletResponse response, ExcelType excelType, int windowSize) {
        ExcelWriterContext context = new ExcelWriterContext();
        context.setFileName(StringUtils.hasText(fileName) ? fileName : LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
        context.setExcelType(ExcelInitializerMeta.INSTANT.initType(null, ExecMode.WRITE) == null ? excelType : ExcelInitializerMeta.INSTANT.initType(null, ExecMode.WRITE));
        context.setExcelEntity(Object.class);
        context.setBind(false);
        return new ExcelSimpleWriter(context, windowSize, response);
    }

    /**
     * Create an Excel bind reader
     *
     * @param file       Excel file
     * @param excelClass Excel mapped entity
     * @param ignores    The name of the header to be ignored during import.
     *                   If it is the parent header, all children below it will be ignored
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelBindReader<R> createReader(MultipartFile file, Class<R> excelClass, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getOriginalFilename());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createReader(file.getInputStream(), excelClass, excelType, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel bind reader
     *
     * @param file       Excel file
     * @param excelClass Excel mapped entity
     * @param ignores    The name of the header to be ignored during import.
     *                   If it is the parent header, all children below it will be ignored
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelBindReader<R> createReader(File file, Class<R> excelClass, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getName());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createReader(new FileInputStream(file), excelClass, excelType, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel bind reader
     *
     * @param inputStream Excel file inputStream
     * @param excelEntity Excel entity
     * @param ignores     Ignore the array of actual Excel table headers that you read when importing
     * @param excelType   Excel file type
     * @param <R>         Entity type
     * @return ExcelReader
     */
    public static <R> ExcelBindReader<R> createReader(InputStream inputStream, Class<R> excelEntity, ExcelType excelType, String... ignores) {
        Objects.requireNonNull(excelEntity, "Excel mapping class cannot be null");
        Excel excel = excelEntity.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "@Excel annotation was not found on the " + excel);
        ExcelReaderContext<R> readerContext = new ExcelReaderContext<>(excelEntity, BeanUtils.getExcelFieldsMap(excelEntity), ignores);
        readerContext.setUniqueKey("".equals(excel.uniqueKey()) ? excelEntity.getName() : excel.uniqueKey());
        return new ExcelBindReader<>(readerContext, inputStream, excelType, excel.cacheRow(), excel.bufferSize());
    }

    /**
     * Create an Excel simple reader
     *
     * @param file       Excel file
     * @param ignores    Ignore the array of actual Excel table headers that you read when importing
     * @param bufferSize buffer size to use when reading InputStream to file, only XLSX
     * @param cacheRow   How many lines of data in the Excel file need to be saved when imported, only XLSX
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelSimpleReader<R> createSimpleReader(File file, int cacheRow, int bufferSize, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getName());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createSimpleReader(new FileInputStream(file), excelType, cacheRow, bufferSize, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel simple reader
     *
     * @param file    Excel file
     * @param ignores Ignore the array of actual Excel table headers that you read when importing
     * @param <R>     Entity type
     * @return ExcelReader
     */
    public static <R> ExcelSimpleReader<R> createSimpleReader(File file, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getName());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createSimpleReader(new FileInputStream(file), excelType, 100, 2048, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel simple reader
     *
     * @param file       Excel file
     * @param ignores    Ignore the array of actual Excel table headers that you read when importing
     * @param bufferSize buffer size to use when reading InputStream to file, only XLSX
     * @param cacheRow   How many lines of data in the Excel file need to be saved when imported, only XLSX
     * @param <R>        Entity type
     * @return ExcelReader
     */
    public static <R> ExcelSimpleReader<R> createSimpleReader(MultipartFile file, int cacheRow, int bufferSize, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getOriginalFilename());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createSimpleReader(file.getInputStream(), excelType, cacheRow, bufferSize, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel simple reader
     *
     * @param file    Excel file
     * @param ignores Ignore the array of actual Excel table headers that you read when importing
     * @param <R>     Entity type
     * @return ExcelReader
     */
    public static <R> ExcelSimpleReader<R> createSimpleReader(MultipartFile file, String... ignores) {
        try {
            ExcelType excelType = ExcelUtils.getExcelType(file.getOriginalFilename());
            if (excelType == null) {
                throw new ExcelTemplateException("File type does not belong to Excel");
            }
            return createSimpleReader(file.getInputStream(), excelType, 100, 2048, ignores);
        } catch (IOException e) {
            throw new ExcelException("Create excel reader error," + e.getMessage());
        }
    }

    /**
     * Create an Excel simple reader
     *
     * @param inputStream Excel file inputStream
     * @param ignores     Ignore the array of actual Excel table headers that you read when importing
     * @param excelType   Excel file type
     * @param bufferSize  buffer size to use when reading InputStream to file, only XLSX
     * @param cacheRow    How many lines of data in the Excel file need to be saved when imported, only XLSX
     * @param <R>         Entity type
     * @return ExcelReader
     */
    public static <R> ExcelSimpleReader<R> createSimpleReader(InputStream inputStream, ExcelType excelType, int cacheRow, int bufferSize, String... ignores) {
        ExcelReaderContext<R> readerContext = new ExcelReaderContext<>(null, null, ignores);
        return new ExcelSimpleReader<>(readerContext, inputStream, excelType, cacheRow, bufferSize);
    }
}
