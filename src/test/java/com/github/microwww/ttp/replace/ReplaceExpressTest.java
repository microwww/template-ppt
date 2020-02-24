package com.github.microwww.ttp.replace;

import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class ReplaceExpressTest {

    @Test
    public void testReplace() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        XSLFTextBox box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(100, 100, 200, 200));
        int start = box.getTextParagraphs().size();

        box.addNewTextParagraph().addNewTextRun().setText("${will-replaced}");

        XSLFTextParagraph graph = box.addNewTextParagraph();
        graph.addNewTextRun().setText("${will");
        graph.addNewTextRun().setText("-replaced}");

        graph = box.addNewTextParagraph();
        graph.addNewTextRun().setText("${will");
        graph.addNewTextRun().setText("-replaced");
        graph.addNewTextRun().setText("}");

        assertEquals(3, box.getTextParagraphs().size() - start);

        graph = box.addNewTextParagraph();
        graph.addNewTextRun().setText("中国${will");
        graph.addNewTextRun().setText("-replaced");
        graph.addNewTextRun().setText("}人民");

        box.addNewTextParagraph().addNewTextRun().setText("中华${will-replaced}人民");

        for (int i = start; i < box.getTextParagraphs().size(); i++) {
            XSLFTextParagraph gps = box.getTextParagraphs().get(i);
            Tools.replace(gps, "${will-replaced}", "TEST" + i);
            assertTrue(gps.getText().contains("TEST"));
            assertTrue(!gps.getText().contains("replaced"));
        }

        graph = box.addNewTextParagraph();
        graph.addNewTextRun().setText("中国${will");
        graph.addNewTextRun().setText("-replaced");
        graph.addNewTextRun().setText("}人民");

        XSLFTextParagraph copy = Tools.copy(graph);
        Tools.replace(graph, "${will-replaced}", "TEST-last1");
        Tools.replace(copy, "${will-replaced}", "TEST-last2");
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ppt.write(out);
        out.toByteArray();
        ppt = new XMLSlideShow(new ByteArrayInputStream(out.toByteArray()));
        XSLFShape shape = ppt.getSlides().get(0).getShapes().get(0);
        box = (XSLFTextBox) shape;

        for (int i = start; i < box.getTextParagraphs().size(); i++) {
            XSLFTextParagraph gps = box.getTextParagraphs().get(i);
            String text = gps.getText();
            assertTrue(text.contains("TEST"));
            assertTrue(!text.contains("replaced"));
        }

    }
}