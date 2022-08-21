package cn.gjing.excel.executor.write;

import cn.gjing.excel.base.aware.ExcelWorkbookAware;
import cn.gjing.excel.base.aware.ExcelWriteContextAware;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.listener.write.ExcelWriteListener;
import cn.gjing.excel.base.meta.ExcelInitializerMeta;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.util.ListenerChain;
import cn.gjing.excel.executor.util.ParamUtils;
import cn.gjing.excel.executor.write.core.ExcelBaseWriteExecutor;
import cn.gjing.excel.executor.write.core.ExcelClassWriterExecutor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.UUID;

/**
 * Excel base writer
 *
 * @author Gjing
 **/
public abstract class ExcelBaseWriter {
    protected ExcelWriterContext context;
    protected HttpServletResponse response;
    protected ExcelBaseWriteExecutor writeExecutor;
    protected final String defaultSheetName = "Sheet1";

    protected ExcelBaseWriter(ExcelWriterContext context, int windowSize, HttpServletResponse response, ExecMode mode) {
        this.response = response;
        this.context = context;
        ExcelInitializerMeta.INSTANT.initListener(context.getExcelEntity(), mode, context.getListenerCache());
        ExcelType globalType = ExcelInitializerMeta.INSTANT.initType(context.getExcelEntity(), mode);
        if (globalType != null) {
            context.setExcelType(globalType);
        }
        this.chooseResolver(context, windowSize, mode);
        context.getListenerCache().forEach(e -> this.initAware((ExcelWriteListener) e));
    }

    /**
     * Choose resolver
     *
     * @param context    Excel write context
     * @param windowSize Window size, which is flushed to disk when exported
     *                   if the data that has been written out exceeds the specified size
     * @param mode       Write executor mode
     */
    protected void chooseResolver(ExcelWriterContext context, int windowSize, ExecMode mode) {
        switch (this.context.getExcelType()) {
            case XLS:
                context.setWorkbook(new HSSFWorkbook());
                break;
            case XLSX:
                context.setWorkbook(new SXSSFWorkbook(windowSize));
                break;
            default:
        }
        this.writeExecutor = new ExcelClassWriterExecutor(context);
    }

    /**
     * Flush all content to excel of the cache
     */
    public void flush() {
        this.processBind();
        if (ListenerChain.doWorkbookFlushBefore(this.context.getListenerCache(), this.context.getWorkbook())) {
            this.writeExecutor.flush(this.response, this.context);
            if (this.context.getWorkbook() instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) this.context.getWorkbook()).dispose();
            }
            return;
        }
        try {
            if (this.context.getWorkbook() != null) {
                this.context.getWorkbook().close();
            }
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
    }

    /**
     * Flush all content to excel of the cache to local
     *
     * @param path Absolute path to the directory where the file is stored
     */
    public void flushToLocal(String path) {
        this.processBind();
        if (ListenerChain.doWorkbookFlushBefore(this.context.getListenerCache(), this.context.getWorkbook())) {
            this.writeExecutor.flushToLocal(path, this.context);
            if (this.context.getWorkbook() instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) this.context.getWorkbook()).dispose();
            }
            return;
        }
        try {
            if (this.context.getWorkbook() != null) {
                this.context.getWorkbook().close();
            }
        } catch (IOException e) {
            throw new ExcelException(e.getMessage());
        }
    }

    /**
     * Create Excel sheet
     *
     * @param sheetName sheet name
     */
    public void createSheet(String sheetName) {
        Sheet sheet = this.context.getWorkbook().getSheet(sheetName);
        if (sheet != null) {
            this.context.setSheet(sheet);
            return;
        }
        sheet = this.context.getWorkbook().createSheet(sheetName);
        this.context.setSheet(sheet);
        ListenerChain.doCompleteSheet(this.context.getListenerCache(), sheet);
    }

    protected void initAware(ExcelWriteListener excelWriteListener) {
        if (excelWriteListener instanceof ExcelWriteContextAware) {
            ((ExcelWriteContextAware) excelWriteListener).setContext(this.context);
        }
        if (excelWriteListener instanceof ExcelWorkbookAware) {
            ((ExcelWorkbookAware) excelWriteListener).setWorkbook(this.context.getWorkbook());
        }
    }

    private void processBind() {
        if (!this.context.isBind()) {
            return;
        }
        String unqSheetName = "excelUnqSheet";
        Sheet sheet = this.context.getWorkbook().createSheet(unqSheetName);
        sheet.protectSheet(UUID.randomUUID().toString().replaceAll("-", ""));
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(ParamUtils.encodeMd5(this.context.getUniqueKey()));
        this.context.getWorkbook().setSheetHidden(this.context.getWorkbook().getSheetIndex(sheet), true);
    }
}
