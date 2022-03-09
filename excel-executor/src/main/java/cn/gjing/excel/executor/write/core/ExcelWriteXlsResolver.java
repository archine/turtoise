package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;

import java.util.List;
import java.util.Map;

/**
 * XLS resolver
 *
 * @author Gjing
 **/
public class ExcelWriteXlsResolver extends ExcelWriterResolver {

    public ExcelWriteXlsResolver(ExcelWriterContext context, ExecMode execMode) {
        super(context, execMode);
    }

    @Override
    public ExcelWriterResolver writeHead(boolean needHead, Map<String, String[]> dropdownBoxValues) {
        super.writeExecutor.writeHead(needHead, dropdownBoxValues);
        return this;
    }

    @Override
    public void write(List<?> data) {
        super.writeExecutor.writeBody(data);
    }
}
