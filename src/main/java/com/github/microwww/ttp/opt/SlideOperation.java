package com.github.microwww.ttp.opt;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class SlideOperation extends Operation {

    @Override
    public void parse(ParseContext context) {
        XSLFSlide form = context.getTemplateShow().getSlides().get(Integer.valueOf(this.getNode()[0]));
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
