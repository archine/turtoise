package cn.gjing.excel.executor.read;

import cn.gjing.excel.base.annotation.Excel;
import cn.gjing.excel.base.aware.ExcelReaderContextAware;
import cn.gjing.excel.base.aware.ExcelWorkbookAware;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.read.ExcelReadListener;
import cn.gjing.excel.base.meta.ExcelInitializerMeta;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.read.core.ExcelBaseReadExecutor;
import cn.gjing.excel.executor.read.core.ExcelClassReadExecutor;
import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * Excel base reader
 *
 * @author Gjing
 **/
public abstract class ExcelBaseReader<R> {
    protected ExcelReaderContext<R> context;
    protected InputStream inputStream;
    protected ExcelBaseReadExecutor<R> baseReadExecutor;
    protected final String defaultSheetName = "Sheet1";

    public ExcelBaseReader(ExcelReaderContext<R> context, InputStream inputStream, ExcelType excelType, Excel excel, ExecMode execMode) {
        this.context = context;
        this.inputStream = inputStream;
        this.chooseResolver(excelType, excel);
        ExcelInitializerMeta.INSTANT.initListener(context.getExcelEntity(), execMode, context.getListenerCache());
    }

    /**
     * Release resources after the read is complete
     */
    public void finish() {
        try {
            if (this.inputStream != null) {
                this.inputStream.close();
            }
            if (this.context.getWorkbook() != null) {
                this.context.getWorkbook().close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void chooseResolver(ExcelType excelType, Excel excel) {
        switch (excelType) {
            case XLS -> {
                try {
                    this.context.setWorkbook(new HSSFWorkbook(this.inputStream));
                } catch (NotOLE2FileException | OfficeXmlFileException e) {
                    e.printStackTrace();
                    throw new ExcelTemplateException();
                } catch (IOException e) {
                    throw new ExcelException("Init workbook error, " + e.getMessage());
                }
            }
            case XLSX -> {
                Workbook workbook;
                try {
                    workbook = StreamingReader.builder()
                            .rowCacheSize(excel.cacheRow())
                            .bufferSize(excel.bufferSize())
                            .setReadShapes(excel.shape())
                            .setReadHyperlinks(excel.hyperlink())
                            .open(this.inputStream);
                } catch (NotOfficeXmlFileException e) {
                    this.finish();
                    e.printStackTrace();
                    throw new ExcelTemplateException();
                }
                this.context.setWorkbook(workbook);
            }
            default -> throw new ExcelException("Excel type invalid");
        }
        this.baseReadExecutor = new ExcelClassReadExecutor<>(this.context);
    }

    @SuppressWarnings("unchecked")
    protected void initAware(ExcelReadListener excelReadListener) {
        if (excelReadListener instanceof ExcelReaderContextAware) {
            ((ExcelReaderContextAware<R>) excelReadListener).setContext(this.context);
        }
        if (excelReadListener instanceof ExcelWorkbookAware) {
            ((ExcelWorkbookAware) excelReadListener).setWorkbook(this.context.getWorkbook());
        }
    }
}
