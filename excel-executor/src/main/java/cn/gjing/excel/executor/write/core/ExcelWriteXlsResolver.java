package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.meta.ExecMode;

import java.util.List;

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
    public ExcelWriterResolver writeHead(boolean needHead) {
        super.writeExecutor.writeHead(needHead);
        return this;
    }

    @Override
    public void write(List<?> data) {
        super.writeExecutor.writeBody(data);
    }
}
