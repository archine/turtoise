package cn.gjing.excel.executor.read.core;

import cn.gjing.excel.base.aware.ExcelWorkbookAware;
import cn.gjing.excel.base.context.ExcelReaderContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.exception.ExcelTemplateException;
import cn.gjing.excel.base.listener.read.ExcelReadListener;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.meta.InitializerMeta;
import cn.gjing.excel.executor.read.aware.ExcelReaderContextAware;
import com.monitorjbl.xlsx.StreamingReader;
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

    public ExcelBaseReader(ExcelReaderContext<R> context, InputStream inputStream, ExcelType excelType, int cacheRowSize, int bufferSize, ExecMode execMode) {
        this.context = context;
        this.inputStream = inputStream;
        this.chooseResolver(excelType, cacheRowSize, bufferSize, execMode);
        InitializerMeta.INSTANT.init(context.getExcelEntity(), ExecMode.READ, context.getListenerCache());
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

    private void chooseResolver(ExcelType excelType, int cacheRowSize, int bufferSize, ExecMode execMode) {
        switch (excelType) {
            case XLS:
                try {
                    this.context.setWorkbook(new HSSFWorkbook(this.inputStream));
                } catch (NotOLE2FileException | OfficeXmlFileException exception) {
                    exception.printStackTrace();
                    throw new ExcelTemplateException();
                } catch (IOException e) {
                    throw new ExcelException("Init workbook error, " + e.getMessage());
                }
                break;
            case XLSX:
                Workbook workbook;
                try {
                    workbook = StreamingReader.builder().rowCacheSize(cacheRowSize).bufferSize(bufferSize).open(this.inputStream);
                } catch (NotOfficeXmlFileException e) {
                    e.printStackTrace();
                    throw new ExcelTemplateException();
                }
                this.context.setWorkbook(workbook);
                break;
            default:
                throw new ExcelException("Excel type cannot be null");
        }
        this.baseReadExecutor = execMode == ExecMode.BIND ? new ExcelBindReadExecutor<>(this.context) : new ExcelSimpleReadExecutor<>(this.context);
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
