package com.github.microwww.ttp;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel._HelpTest;
import org.junit.Assert;
import org.junit.Test;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.UUID;

public class ToolsTest {
    public static final String PATH = _HelpTest.PATH;

    @Test
    public void testDate1904() throws IOException, InvalidFormatException {
        XMLSlideShow target, template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "chart.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "chart.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }
        // chart.pptx , time is 2002, importContent time is 2006 ! set date1904, chart.zip/ppt/charts/chart1.xml : <c:date1904 val="0"/>
        target.createSlide().importContent(template.getSlides().get(0));
        // No working
        Tools.findChart(target.getSlides().get(0), 0).getWorkbook().getCTWorkbook().getWorkbookPr().setDate1904(true);

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }

    @Test
    public void testPosition() throws IOException {
        XMLSlideShow target, template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }
        XSLFSlide from = template.getSlides().get(1);
        XSLFTable shape = (XSLFTable) from.getShapes().get(1);

        XSLFSlide slide = target.createSlide();
        Tools.copyChart(from, 0, slide);

        Tools.copyTable(slide, shape);
        XSLFTable nt1 = Tools.copyTable(slide, shape);
        double dy = nt1.getAnchor().getHeight() * 1.5;
        Tools.copyChart(from, 0, slide, new Rectangle2D.Double(0, dy, 0, 0));

        Rectangle2D npos = Tools.delta(nt1.getAnchor(), 0, dy, 0, 0);
        nt1.setAnchor(npos);

        File file = new File(PATH, UUID.randomUUID().toString() + ".pptx");
        target.write(new FileOutputStream(file));

        try (FileInputStream in = new FileInputStream(file)) {
            target = new XMLSlideShow(in);
        }

        slide = target.getSlides().get(0); //
        Assert.assertEquals(4, slide.getShapes().size());
    }
}