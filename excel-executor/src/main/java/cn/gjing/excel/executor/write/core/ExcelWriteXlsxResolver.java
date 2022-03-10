package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.context.ExcelWriterContext;
import cn.gjing.excel.base.meta.ExecMode;

import java.util.List;

/**
 * Xlsx resolver
 *
 * @author Gjing
 **/
public class ExcelWriteXlsxResolver extends ExcelWriterResolver {

    public ExcelWriteXlsxResolver(ExcelWriterContext context, ExecMode execMode) {
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
