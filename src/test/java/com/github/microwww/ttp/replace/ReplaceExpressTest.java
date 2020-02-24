package com.github.microwww.ttp.replace;

import com.github.microwww.ttp.Assert;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

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
            replace(gps, "${will-replaced}", "TEST" + i);
            assertTrue(gps.getText().contains("TEST"));
            assertTrue(!gps.getText().contains("replaced"));
        }

        graph = box.addNewTextParagraph();
        graph.addNewTextRun().setText("中国${will");
        graph.addNewTextRun().setText("-replaced");
        graph.addNewTextRun().setText("}人民");

        XSLFTextParagraph copy = copy(graph);
        replace(graph, "${will-replaced}", "TEST-last1");
        replace(copy, "${will-replaced}", "TEST-last2");
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

    public static void replace(XSLFTextParagraph paragraph, String origin, String replacement) {
        CTTextParagraph xml = paragraph.getXmlObject();
        StringBuffer buffer = new StringBuffer();
        List<CTRegularTextRun> runs = xml.getRList();
        for (CTRegularTextRun rText : runs) {
            buffer.append(rText.getT());
        }
        int idx = buffer.indexOf(origin), fromRun = -1, toRun = -1, cursor = 0;
        int end = idx + origin.length();
        for (int i = 0; i < runs.size(); i++) {
            CTRegularTextRun rText = runs.get(i);
            String text = rText.getT(), start = "", fix = "";
            int len = text.length();

            if (idx >= cursor + len) {// 未开始
                cursor += len;
                continue;
            }
            if (end < cursor) { // 已结束
                break;
            }

            if (idx >= cursor && idx <= cursor + len) {// 起始
                fromRun = i;
                // int fromRunPosition = idx - cursor;
                start = text.substring(0, idx - cursor) + replacement;// 开始替换:取前半段然后追加替换字符
            }
            if (end >= cursor && end <= cursor + len) {
                toRun = i;
                // int toRunPosition = end - cursor;
                fix = text.substring(end - cursor); // 取后半段
            }
            rText.setT(start + fix);
            cursor += len;
        }
        Assert.isTrue(fromRun >= 0 && toRun >= 0, "NOT find message by XML : " + origin + ", check it");
    }

    public static XSLFTextParagraph copy(XSLFTextParagraph paragraph) {
        XSLFTextShape shape = paragraph.getParentShape();
        XSLFTextParagraph newPg = shape.addNewTextParagraph();
        newPg.getXmlObject().set(paragraph.getXmlObject().copy());
        return newPg;
    }
}