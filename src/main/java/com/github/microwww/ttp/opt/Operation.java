package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public abstract class Operation {

    private String[] expresses;
    private String[] params;

    public abstract void parse(XSLFSheet slide, List<Operation> parsed);

        public String[] getExpresses() {
        return expresses;
    }

    public void setExpresses(String[] expresses) {
        this.expresses = expresses;
    }

    public String[] getParams() {
        return params;
    }

    public void setParams(String[] params) {
        this.params = params;
    }

    public List<?> searchElement(XSLFSheet slide, int from) throws ClassNotFoundException {
        String[] exp = getExpresses();
        Assert.isTrue((exp.length - from) % 2 == 0, "express message pare with shape / index !");
        int level = from;
        List<Object> parent = new ArrayList<>();
        if (level < exp.length) {// 第一级
            String cname = "org.apache.poi.xslf.usermodel." + exp[level];
            String index = exp[level + 1];
            List<Range> list = Operation.searchRanges(index);
            if (Class.forName(cname).equals(XSLFChart.class)) {
                List<XSLFGraphicChart> charts = _Help.listCharts(slide);
                for (int i = 0; i < charts.size(); i++) {
                    for (Range rg : list) {
                        if (rg.isIn(i)) {
                            parent.add(charts.get(i));
                            break;
                        }
                    }
                }
            } else {
                int idx = 0;
                for (XSLFShape shape : slide.getShapes()) {
                    if (Class.forName(cname).isInstance(shape)) {
                        for (Range rg : list) {
                            if (rg.isIn(idx)) {
                                parent.add(shape);
                                idx++;
                                break;
                            }
                        }
                    }
                }
            }
        }
        level += 2;
        if (level < exp.length) {// 第二级
            List<Object> next = new ArrayList<>();
            String cname = "org.apache.poi.xslf.usermodel." + exp[level];
            String index = exp[level + 1];
            List<Range> list = Operation.searchRanges(index);
            if (Class.forName(cname).equals(XSLFTableRow.class)) {
                for (Object cn : parent) {
                    if (cn instanceof XSLFTable) {
                        List<XSLFTableRow> rows = ((XSLFTable) cn).getRows();
                        for (int i = 0; i < rows.size(); i++) {
                            for (Range r : list) {
                                if (r.isIn(i)) {
                                    next.add(rows.get(i));
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            parent = next;
        }
        level += 2;
        List<Object> three = new ArrayList<>();
        if (level < exp.length) {// 第二级
            String cname = "org.apache.poi.xslf.usermodel." + exp[level];
            String index = exp[level + 1];
            List<Range> list = Operation.searchRanges(index);
            if (Class.forName(cname).equals(XSLFTableRow.class)) {
                for (Object cn : parent) {
                    if (cn instanceof XSLFTableRow) {
                        List<XSLFTableCell> cells = ((XSLFTableRow) cn).getCells();
                        for (int i = 0; i < cells.size(); i++) {
                            for (Range r : list) {
                                if (r.isIn(i)) {
                                    three.add(cells.get(i));
                                }
                            }
                        }
                    }
                }
            }
            parent = three;
        }
        return parent;
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
