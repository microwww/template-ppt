package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel._HelpTest;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import static org.junit.Assert.*;

public class ParseExpressesTest {

    @Test
    public void parse() throws IOException {
        ParseExpresses exp = new ParseExpresses(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        exp.parse();
        assertEquals(14, exp.getOperations().size());
    }

    @Test
    public void testFile() throws IOException {
        ParseExpresses exp = new ParseExpresses(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        exp.parse();
        List<Operation> ops = exp.getOperations();
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        XSLFSlide slide = template.getSlides().get(0);
        context.setTemplate(slide);
        for (int i = 0; i < ops.size(); i++) {
            Operation o = ops.get(i);
            o.parse(context);
        }
    }
}