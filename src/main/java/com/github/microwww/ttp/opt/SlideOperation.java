package com.github.microwww.ttp.opt;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.util.List;

public class SlideOperation extends Operation {

    private static List<XSLFSlide> origins = null;

    @Override
    public void parse(ParseContext context) {
        if (origins == null) {
            origins = context.getTemplateShow().getSlides();
        }
        XSLFSlide form = origins.get(Integer.valueOf(this.getNode()[0]));
        context.setTemplate(form);
    }

    @Override
    public void setNode(String[] node) {
        if (node.length > 0 && node[node.length - 1].equals("]")) {
            node = ArrayUtils.subarray(node, 0, node.length - 1);
        }
        super.setNode(node);
    }
}
