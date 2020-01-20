package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.util.List;

public class SlideOperation extends Operation {

    @Override
    public void parse(XSLFSheet slide, List<Operation> ops) {
        XSLFSlide form = slide.getSlideShow().getSlides().get(Integer.valueOf(this.getExpresses()[2]));
        boolean start = false;
        SlideOperation next = null;
        for (int i = 0; i < ops.size(); i++) {
            Operation o = ops.get(i);
            if (!start) {
                if (o.equals(this)) {
                    start = true;
                }
                continue;
            }
            if (this.getClass().isInstance(o)) {
                o.parse(slide, ops); // ROOOOOOOLL
                break;
            }
            o.parse(form, ops);
        }
    }
}
