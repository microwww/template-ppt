package org.apache.poi.xslf.usermodel;

import com.github.microwww.ttp.util._Help;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.UUID;

public class _HelpTest {

    public static final String PATH = _HelpTest.class.getResource("/").getFile();

    @Test
    public void copyTable() throws IOException {
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
        XSLFSlide slide = target.createSlide();
        XSLFTable shape = (XSLFTable) template.getSlides().get(0).getShapes().get(0);
        // 1
        XSLFTable table = _Help.copyTable(slide, shape);
        Assert.assertEquals(table.getRows().size(), shape.getRows().size());
        Assert.assertEquals(table.getRows().get(0).getCells().size(), shape.getRows().get(0).getCells().size());

        // 2
        XSLFTableRow orow = shape.getRows().get(1);
        XSLFTableRow row = _Help.copyTableRow(table, orow);
        Assert.assertEquals(row.getCells().size(), orow.getCells().size());

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }

    @Test
    public void copyChart() throws IOException, InvalidFormatException {
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

        target.createSlide().importContent(template.getSlides().get(0));

        XSLFSlide slide = target.createSlide();
        XSLFChart chart = _Help.copyChart(template.getSlides().get(0), 0, slide, (c, val) -> {
            slide.addChart(c, _Help.rectanglePx2point(val.getGraphic().getAnchor(), 0, 0, 0, 0));
        });
        Assert.assertNotNull(chart);

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }
}