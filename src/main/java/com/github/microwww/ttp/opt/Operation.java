package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.beanutils.MethodUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public abstract class Operation {

    private static final Logger logger = LoggerFactory.getLogger(Operation.class);

    private String prefix;
    private String[] node;
    private String[] params;

    public abstract void parse(ParseContext context);

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

    public List<?> search(ParseContext context) {
        String[] exp = getNode();
        Assert.isTrue(exp.length % 2 == 0, "express message pare with shape / index !");

        List<Object> content = Collections.singletonList(context.getTemplateShow());
        for (int i = 0; i < exp.length; i += 2) {
            List<Object> next = new ArrayList<>();
            for (Object cnt : content) {
                List<Object> list = searchElement(context, cnt, exp[i], exp[i + 1]);
                next.addAll(list);
            }
            content = next;
        }
        return content;
    }

    private List<Object> searchElement(ParseContext context, Object content, String exp, String range) {
        try {
            Object element = MethodUtils.invokeMethod(this, "findElement",
                    new Object[]{context, content, exp, range});
            return (List<Object>) element;
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
                for (XSLFShape shape : slide.getShapes()) {
                    if (clazz.isInstance(shape)) {
                        for (Range rg : list) {
                            if (rg.isIn(idx)) {
                                res.add(shape);
                                idx++;
                                break;
                            }
                        }
                    }
                }
            } catch (ClassNotFoundException e) {
                throw new RuntimeException("Exception not support ! Must in package: org.apache.poi.xslf.usermodel", e);
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
                this.findElement(context, row, exp, range);
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
     * @param index
     * @return
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
}
