package cn.gjing.excel.executor.write.core;

import cn.gjing.excel.base.meta.ExecMode;
import cn.gjing.excel.executor.write.context.ExcelWriterContext;

import java.util.List;
import java.util.Map;

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
    public ExcelWriterResolver writeHead(boolean needHead, Map<String, String[]> boxValues) {
        super.writeExecutor.writeHead(needHead, boxValues);
        return this;
    }

    @Override
    public void write(List<?> data) {
        super.writeExecutor.writeBody(data);
    }
}
