package com.github.microwww.ttp;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel._HelpTest;
import org.junit.Test;

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
}