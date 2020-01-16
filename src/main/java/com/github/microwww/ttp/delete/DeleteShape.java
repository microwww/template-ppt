package com.github.microwww.ttp.delete;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.util.List;

public class DeleteShape {
    public void delete(XSLFShape shape) {
        shape.getSheet().removeShape(shape);
    }

    public void delete(XSLFSheet sheet) {
        List<XSLFSlide> slides = sheet.getSlideShow().getSlides();
        for (int i = slides.size() - 1; i >= 0; i--) {
            if (slides.get(i) == sheet) {
                sheet.getSlideShow().removeSlide(i);
            }
        }
    }
}
