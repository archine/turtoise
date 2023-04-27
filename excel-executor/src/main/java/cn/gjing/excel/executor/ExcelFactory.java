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
import cn.gjing.excel.executor.write.ExcelBindWriter;
import cn.gjing.excel.executor.write.ExcelSimpleWriter;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.WritableByteChannel;
import java.nio.file.Files;
import java.nio.file.Paths;
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
    public static ExcelBindWriter createWriter(Class<?> excelEntity, HttpServletResponse response, String... ignores) {
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
    public static ExcelBindWriter createWriter(String fileName, Class<?> excelEntity, HttpServletResponse response, String... ignores) {
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
        context.setIdCard("".equals(excel.idCard()) ? excelEntity.getName() : excel.idCard());
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
    public static ExcelSimpleWriter createSimpleWriter(String fileName, HttpServletResponse response, ExcelType excelType, int windowSize) {
        ExcelWriterContext context = new ExcelWriterContext();
        context.setFileName(StringUtils.hasText(fileName) ? fileName : LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
        context.setExcelEntity(null);
        context.setExcelType(excelType);
        context.setBind(false);
        context.setHeaderSeries(0);
        return new ExcelSimpleWriter(context, windowSize, response);
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
                throw new ExcelTemplateException("file type must be excel");
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
                throw new ExcelTemplateException("file type must be excel");
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
     * @param excelClass  Object class to be generated
     * @param excelType   Excel file type
     * @param <R>         Entity type
     * @return ExcelReader
     */
    public static <R> ExcelClassReader<R> createReader(InputStream inputStream, Class<R> excelClass, ExcelType excelType) {
        Objects.requireNonNull(excelClass, "Excel mapping class cannot be null");
        Excel excel = excelClass.getAnnotation(Excel.class);
        Objects.requireNonNull(excel, "@Excel annotation was not found on the " + excel);
        ExcelReaderContext<R> readerContext = new ExcelReaderContext<>(excelClass);
        readerContext.setIdCard("".equals(excel.idCard()) ? excelClass.getName() : excel.idCard());
        readerContext.setFieldProperties(BeanUtils.getExcelFiledProperties(excelClass, null));
        return new ExcelClassReader<>(readerContext, inputStream, excelType, excel);
    }

    /**
     * Output the small file to the network for download
     * Note: Not suitable for transferring large files (GB level)
     *
     * @param file     File
     * @param filename Downloaded file name
     * @param response HttpServletResponse
     */
    public static void transferToNetwork(File file, String filename, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Type", Files.probeContentType(Paths.get(file.getAbsolutePath())));
        response.setContentLength((int) file.length());
        String encodeFileName = URLEncoder.encode(filename, "utf-8").replaceAll("\\+", "%20");
        String dispositionVal = "attachment; filename=" + encodeFileName + ";" + "filename*=" + "utf-8''" + encodeFileName;
        response.setHeader("Content-disposition", dispositionVal);
        try (FileInputStream fileInputStream = new FileInputStream(file);
             FileChannel fileChannel = fileInputStream.getChannel();
             OutputStream outputStream = response.getOutputStream();
             WritableByteChannel writableByteChannel = Channels.newChannel(outputStream)) {
            fileChannel.transferTo(0, fileChannel.size(), writableByteChannel);
        }
    }
}
