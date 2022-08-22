package cn.gjing.excel.base.meta;

import org.springframework.expression.EvaluationContext;
import org.springframework.expression.spel.standard.SpelExpressionParser;

/**
 * EL global expression parser
 *
 * @author Gjing
 **/
public enum ELMeta {
    PARSER;

    private final SpelExpressionParser parser = new SpelExpressionParser();

    public SpelExpressionParser getParser() {
        return PARSER.parser;
    }

    /**
     * El expression parsing
     *
     * @param expr       el expression
     * @param context    EvaluationContext
     * @param returnType Return value generic
     * @param <R>        R
     * @return R
     */
    public <R> R parse(String expr, EvaluationContext context, Class<R> returnType) {
        return parser.parseExpression(expr).getValue(context, returnType);
    }

    /**
     * El expression parsing
     *
     * @param expr    el expression
     * @param context EvaluationContext
     * @return Obj
     */
    public Object parse(String expr, EvaluationContext context) {
        return parser.parseExpression(expr).getValue(context);
    }
}
