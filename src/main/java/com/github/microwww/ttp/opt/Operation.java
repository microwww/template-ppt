package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSheet;

import java.util.ArrayList;
import java.util.List;

public class Operation {

    private String[] expresses;
    private String[] params;

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

    public void getElement(XSLFSheet slide, int from) throws ClassNotFoundException {
        String[] exp = getExpresses();
        Assert.isTrue((exp.length - from) % 2 == 0, "express message pare with shape / index !");
        for (int i = from; i < exp.length; i += 2) {
            String cname = "org.apache.poi.xslf.usermodel." + exp[i];
            String index = exp[i + 1];
            List<Range> list = Operation.searchRanges(index);
            int idx = 0;
            List<XSLFShape> inone = new ArrayList<>();
            for (XSLFShape shape : slide.getShapes()) {
                if (Class.forName(cname).isInstance(shape)) {
                    for (Range rg : list) {
                        if (rg.isIn(idx)) {
                            inone.add(shape);
                        }
                    }
                }
            }
        }
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
