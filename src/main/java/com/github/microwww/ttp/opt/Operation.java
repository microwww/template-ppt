package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.util.DataUtil;
import com.github.microwww.ttp.util.DefaultMemberAccess;
import com.github.microwww.ttp.util._Help;
import ognl.*;
import org.apache.commons.beanutils.MethodUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Stack;

public abstract class Operation {

    private static final Logger logger = LoggerFactory.getLogger(Operation.class);

    private static final OgnlContext context = new OgnlContext(new DefaultClassResolver(), new DefaultTypeConverter(),
            new DefaultMemberAccess(true));

    private String prefix;
    private String[] node;
    private String[] params;

    final protected List<Operation> childrenOperations = new ArrayList<>();
    protected Operation parentOperations;

    public abstract void parse(ParseContext context);

    public void parse(ParseContext context, String method) {
        Stack<Object> container = context.getContainer();
        List<Object> origin = new ArrayList<>(container);
        try {
            List<Stack<Object>> search = this.searchStack(context);
            for (Stack<Object> item : search) {
                container.clear();
                container.addAll(item);
                thisInvoke(method, context, item.peek());
            }
        } finally {
            container.clear();
            container.addAll(origin);
        }
    }

    public String[] getNode() {
        return node;
    }

    public void setNode(String[] node) {
        this.node = node;
    }

    public String[] getParams() {
        return params;
    }

    public void setParams(String[] params) {
        this.params = params;
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public Operation getParentOperations() {
        return parentOperations;
    }

    public void setParentOperations(Operation parentOperations) {
        this.parentOperations = parentOperations;
    }

    public void addChildrenOperation(Operation operation) {
        this.childrenOperations.add(operation);
    }

    public List<?> search(ParseContext context) {
        List<Stack<Object>> stacks = this.searchStack(context);
        List<Object> res = new ArrayList<>();
        for (Stack<Object> stack : stacks) {
            res.add(stack.peek());
        }
        return res;
    }

    public List<Stack<Object>> searchStack(ParseContext context) {
        String[] exp = getNode();
        Assert.isTrue(exp.length % 2 == 0, "express message pare with shape / index !");
        Stack<Object> stack = new Stack<>();
        stack.addAll(context.getContainer());
        List<Stack<Object>> content = Collections.singletonList(stack);
        for (int i = 0; i < exp.length; i += 2) {
            List<Stack<Object>> next = new ArrayList<>();
            //nodeStack.push(next);
            for (Stack<Object> cnt : content) {
                List<Object> list = searchElement(context, cnt.peek(), exp[i], exp[i + 1]);
                // next.addAll(list);
                for (Object last : list) {
                    Stack st = new Stack<>();
                    st.addAll(cnt);
                    st.push(last);
                    next.add(st);
                }
            }
            content = next;
        }
        return content;
    }

    public <T> T getValue(String express, Stack<Object> models, Class<T> clazz) {
        tryParent(models);
        for (int i = models.size(); i > 0; i--) {
            try {
                return (T) Ognl.getValue(express, context, models.get(i - 1), clazz);
            } catch (OgnlException e) {// ignore
            }
        }
        throw new RuntimeException("OGNL express error : " + express);
    }

    public List getCollectionValue(String express, Stack<Object> model) {
        Object value = this.getValue(express, model);
        if (value == null) {
            throw new RuntimeException("OGNL Express value is null, NOT list/array");
        }
        return DataUtil.toList(value);
    }

    public Object getValue(String express, Stack<Object> models) {
        tryParent(models);
        for (int i = models.size(); i > 0; i--) {
            try {
                return Ognl.getValue(express, context, models.get(i - 1));
            } catch (OgnlException e) {// ignore
                logger.debug("Try OGNL error : {}", express, e);
            }
        }
        throw new RuntimeException("OGNL express error : " + express + "."
                + " SET logger : com.github.microwww.ttp.opt.Operation level to DEBUG , see more information");
    }

    public void tryParent(Stack<Object> models) {
        Object next = null;
        for (int k = 0; k < models.size(); k++) {
            Object md = models.get(k);
            if (md instanceof RepeatDomain) {
                ((RepeatDomain) md).setParent(next);
            }
            next = models.get(k);
        }
    }

    private List<Object> searchElement(ParseContext context, Object content, String exp, String range) {
        Object element = thisInvoke("findElement", new Object[]{context, content, exp, range});
        return (List<Object>) element;
    }

    protected Object thisInvoke(String method, Object... params) {
        try {
            if (params == null) {
                params = new Object[]{};
            }
            return MethodUtils.invokeMethod(this, method, params);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    // default 默认方法
    public List<Object> findElement(ParseContext context, Object content, String exp, String range) {
        logger.warn("Skip PPT express {} in {}", exp, content.getClass());
        return Collections.emptyList();
    }

    public List<Object> findElement(ParseContext context, XMLSlideShow content, String exp, String range) {
        List<Object> res = new ArrayList<>();
        List<Range> rgs = Operation.searchRanges(range);
        if (XSLFSlide.class.getSimpleName().equals(exp)) {
            List<XSLFSlide> list = content.getSlides();
            for (int i = 0; i < list.size(); i++) {
                for (Range r : rgs) {
                    if (r.isIn(i)) {
                        res.add(list.get(i));
                    }
                }
            }
        }
        return res;
    }

    // 一级
    public List<Object> findElement(ParseContext context, XSLFSheet slide, String exp, String range) {
        List<Range> list = Operation.searchRanges(range);
        List<Object> res = new ArrayList<>();
        if (XSLFChart.class.getSimpleName().equals(exp)) {
            List<XSLFGraphicChart> charts = _Help.listCharts(slide);
            for (int i = 0; i < charts.size(); i++) {
                for (Range rg : list) {
                    if (rg.isIn(i)) {
                        res.add(charts.get(i));
                        break;
                    }
                }
            }
        } else {
            int idx = 0;
            String cname = "org.apache.poi.xslf.usermodel." + exp;
            try {
                Class clazz = Class.forName(cname);
                if (XSLFSlide.class.equals(clazz)) {
                    for (XSLFSlide shape : slide.getSlideShow().getSlides()) {
                        for (Range rg : list) {
                            if (rg.isIn(idx)) {
                                res.add(shape);
                                break;
                            }
                        }
                        idx++;
                    }
                } else
                    for (XSLFShape shape : slide.getShapes()) {
                        if (clazz.isInstance(shape)) {
                            for (Range rg : list) {
                                if (rg.isIn(idx)) {
                                    res.add(shape);
                                    break;
                                }
                            }
                            idx++;
                        }
                    }
            } catch (ClassNotFoundException e) {
                throw new RuntimeException("Exception not support ! Must in package: org.apache.poi.xslf.usermodel", e);
            }
        }
        return res;
    }

    // 二级
    public List<Object> findElement(ParseContext context, XSLFTextShape content, String exp, String range) {
        List<Object> res = new ArrayList<>();
        List<Range> list = Operation.searchRanges(range);
        if (XSLFTextParagraph.class.getSimpleName().equals(exp)) {
            List<XSLFTextParagraph> rows = content.getTextParagraphs();
            for (int i = 0; i < rows.size(); i++) {
                for (Range r : list) {
                    if (r.isIn(i)) {
                        res.add(rows.get(i));
                        break;
                    }
                }
            }
        }
        return res;
    }

    // 二级
    public List<Object> findElement(ParseContext context, XSLFTable content, String exp, String range) {
        List<Object> res = new ArrayList<>();
        List<Range> list = Operation.searchRanges(range);
        if (XSLFTableRow.class.getSimpleName().equals(exp)) {
            List<XSLFTableRow> rows = content.getRows();
            for (int i = 0; i < rows.size(); i++) {
                for (Range r : list) {
                    if (r.isIn(i)) {
                        res.add(rows.get(i));
                        break;
                    }
                }
            }
        } else if (XSLFTableCell.class.getSimpleName().equals(exp)) {
            for (XSLFTableRow row : content.getRows()) {
                List<Object> lise = this.findElement(context, row, exp, range);
                res.addAll(lise);
            }
        }
        return res;
    }

    // 三级
    public List<Object> findElement(ParseContext context, XSLFTableRow content, String exp, String range) {
        List<Object> res = new ArrayList<>();
        List<Range> list = Operation.searchRanges(range);
        if (XSLFTableCell.class.getSimpleName().equals(exp)) {
            List<XSLFTableCell> cells = content.getCells();
            for (int i = 0; i < cells.size(); i++) {
                for (Range r : list) {
                    if (r.isIn(i)) {
                        res.add(cells.get(i));
                    }
                }
            }
        }
        return res;
    }

    /**
     * support three format : 0,2,3 / 0-2,2-4 / 1-4,8-
     *
     * @param index format string
     * @return range
     */
    public static List<Range> searchRanges(String index) {
        List<Range> list = new ArrayList<>();
        String[] ors = StringUtils.split(index, ',');
        for (String or : ors) {
            if (or.indexOf('-') >= 0) {
                String[] fromTo = StringUtils.split(or, '-');
                int to = Integer.MAX_VALUE;
                if (fromTo.length > 1) { // :: 0-
                    to = Integer.valueOf(fromTo[1]);
                }
                Range rg = new Range(Integer.valueOf(fromTo[0]), to);
                list.add(rg);
            } else {
                if (StringUtils.isBlank(or)) {
                    throw new IllegalArgumentException("',' left/right must have number!");
                }
                int r = Integer.valueOf(or);
                Range rg = new Range(r, r + 1);
                list.add(rg);
            }
        }
        return list;
    }

    public ParamMessage[] getParamsWithPattern() {
        String[] params = getParams();
        ParamMessage[] res = new ParamMessage[params.length];
        for (int i = 0; i < params.length; i++) {
            String param = params[i];
            int idx = param.lastIndexOf('|');
            if (idx > 0) {
                res[i] = new ParamMessage(param.substring(0, idx), param.substring(idx + 1));
            } else {
                res[i] = new ParamMessage(param, "{0}");
            }
        }
        return res;
    }

    public static class ParamMessage {
        private final String param;
        private final String pattern;

        public ParamMessage(String param, String pattern) {
            this.param = param;
            this.pattern = pattern;
        }

        public String getParam() {
            return param;
        }

        public String getPattern() {
            return pattern;
        }

        public String format(Object... arg) {
            return MessageFormat.format(this.pattern, arg);
        }
    }
}
