package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.BigTitle;
import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.exception.ExcelException;
import cn.gjing.excel.base.meta.ExcelType;
import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.base.util.ExcelUtils;
import cn.gjing.excel.base.util.ListenerChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;

/**
 * Excel writer resolver
 *
 * @author Gjing
 **/
public abstract class ExcelWriterResolver {
    protected final ExcelWriterContext context;
    protected final ExcelBaseWriteExecutor writeExecutor;

    public ExcelWriterResolver(ExcelWriterContext context, ExecMode mode) {
        this.context = context;
        if (mode == ExecMode.BIND_WRITE) {
            this.writeExecutor = new ExcelBindWriterExecutor(context);
        } else {
            this.writeExecutor = new ExcelSimpleWriterExecutor(context);
        }
    }

    /**
     * Write excel big title
     *
     * @param bigTitle Excel big title
     */
    public void writeTitle(BigTitle bigTitle) {
        if (bigTitle.getLastCol() < 1) {
            bigTitle.setLastCol(this.context.getFieldProperties().size() - 1);
        }
        if (bigTitle.getRowNum() < 1) {
            bigTitle.setRowNum(1);
        }
        if (bigTitle.getFirstCol() < 0) {
            bigTitle.setFirstCol(0);
        }
        if (bigTitle.getRowNum() == 1 && bigTitle.getFirstCol() == bigTitle.getLastCol()) {
            throw new ExcelException("Merged region must contain 2 or more cells");
        }
        int startOffset = bigTitle.getFirstRow() == -1 ? this.context.getSheet().getPhysicalNumberOfRows() : bigTitle.getFirstRow();
        int endOffset = startOffset + bigTitle.getRowNum() - 1;
        Row row;
        for (int i = 0; i < bigTitle.getRowNum(); i++) {
            row = this.context.getSheet().getRow(startOffset + i);
            if (row == null) {
                row = this.context.getSheet().createRow(startOffset + i);
                row.setHeight(bigTitle.getRowHeight());
            }
            Cell cell = row.createCell(bigTitle.getFirstCol());
            ExcelUtils.setCellValue(cell, bigTitle.getContent());
            if (i == 0) {
                ListenerChain.doSetTitleStyle(this.context.getListenerCache(), bigTitle, cell);
            }
        }
        this.context.getSheet().addMergedRegion(new CellRangeAddress(startOffset, endOffset, bigTitle.getFirstCol(), bigTitle.getLastCol()));
    }

    /**
     * Write excel body
     *
     * @param data data
     */
    public abstract void write(List<?> data);

    /**
     * Write excel header
     *
     * @param needHead Is needHead excel entity or sheet?
     * @return this
     */
    public abstract ExcelWriterResolver writeHead(boolean needHead);

    /**
     * Output the contents of the cache
     *
     * @param context  Excel write context
     * @param response response
     */
    public void flush(HttpServletResponse response, ExcelWriterContext context) {
        response.setContentType("application/vnd.ms-excel");
        String fileName = context.getFileName() + (context.getExcelType() == ExcelType.XLS ? ".xls" : ".xlsx");
        OutputStream outputStream = null;
        try {
            String encodeFileName = URLEncoder.encode(fileName, "utf-8").replaceAll("\\+", "%20");
            String dispositionVal = "attachment; filename=" +
                    encodeFileName +
                    ";" +
                    "filename*=" +
                    "utf-8''" +
                    encodeFileName;
            response.setHeader("Content-disposition", dispositionVal);
            outputStream = response.getOutputStream();
            context.getWorkbook().write(outputStream);
        } catch (IOException e) {
            throw new ExcelException("Excel cache data flush failure, " + e.getMessage());
        } finally {
            try {
                if (outputStream != null) {
                    outputStream.flush();
                    outputStream.close();
                }
                if (context.getWorkbook() != null) {
                    context.getWorkbook().close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * Output the contents of the cache to local
     *
     * @param path    Absolute path to the directory where the file is stored
     * @param context Excel write context
     */
    public void flushToLocal(String path, ExcelWriterContext context) {
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream((path.endsWith("/") ? path : path + "/") + context.getFileName() + (context.getExcelType() == ExcelType.XLS ? ".xls" : ".xlsx"));
            context.getWorkbook().write(fileOutputStream);
        } catch (IOException e) {
            throw new ExcelException("Excel cache data flush failure, " + e.getMessage());
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.flush();
                    fileOutputStream.close();
                }
                if (context.getWorkbook() != null) {
                    context.getWorkbook().close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
