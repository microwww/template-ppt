package com.github.microwww.ttp;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class UtilPPTTest {

    // https://github.com/apache/poi/blob/bb2ad49a2fc6c74948f8bb92701807093b525586/src/examples/src/org/apache/poi/xslf/usermodel/ChartFromScratch.java
    @Test
    public void copeTable() throws IOException, InvalidFormatException {
        XMLSlideShow target, template;
        String path = _HelpTest.PATH;
        try (FileInputStream in = new FileInputStream(new File(path, "chart.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(path, "chart.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }

        // copy chart !!!!
        XSLFSlide slide = target.createSlide();
        XSLFChart radar = UtilPPT.findChart(template.getSlides().get(0), 0);
        Rectangle2D anchor = template.getSlides().get(0).getShapes().get(0).getAnchor();

        XSLFChart chart = target.createChart();
        chart.importContent(radar);
        chart.setWorkbook(radar.getWorkbook());
        slide.addChart(chart, UtilPPT.rectanglePx2point(anchor));

        target.write(new FileOutputStream(new File(path, "target.pptx")));

    }

}