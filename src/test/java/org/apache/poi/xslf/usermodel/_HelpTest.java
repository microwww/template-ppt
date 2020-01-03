package org.apache.poi.xslf.usermodel;

import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.UUID;

public class _HelpTest {

    private static final String PATH = _HelpTest.class.getResource("/").getFile();

    @Test
    public void copeTable() throws IOException {
        XMLSlideShow target, ppt;
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            ppt = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }
        XSLFSlide slide = target.createSlide();
        XSLFTable shape = (XSLFTable) ppt.getSlides().get(0).getShapes().get(0);
        // 1
        XSLFTable table = _Help.copeTable(slide, shape);
        Assert.assertEquals(table.getRows().size(), shape.getRows().size());
        Assert.assertEquals(table.getRows().get(0).getCells().size(), shape.getRows().get(0).getCells().size());

        // 2
        XSLFTableRow orow = shape.getRows().get(1);
        XSLFTableRow row = _Help.copyTableRow(table, orow);
        Assert.assertEquals(row.getCells().size(), orow.getCells().size());

        target.write(new FileOutputStream(new File(PATH, UUID.randomUUID().toString() + ".pptx")));
    }
}