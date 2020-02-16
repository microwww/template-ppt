package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel._HelpTest;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;

public class ParseExpressesTest {

    @Test
    public void parse() throws IOException {
        ParseExpresses exp = new ParseExpresses(new File(this.getClass().getResource("/").getFile(), "demo.txt"));
        exp.parse();
        assertEquals(14, exp.getOperations().size());
    }

    @Test
    public void testFile() throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        XSLFSlide slide = template.getSlides().get(0);
        context.setTemplate(slide);

        Map<String, Object> map = new HashMap<>();
        map.put("name", "Hello");
        map.put("age", 100);
        context.setData(map);
        {
            ReplaceOperation rep = new ReplaceOperation();
            rep.setNode(new String[]{"XSLFTable", "0"});
            rep.setParams(new String[]{});
            rep.parse(context);
        }
        {
            ReplaceOperation rep = new ReplaceOperation();
            rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "1", "XSLFTableCell", "1"});
            rep.setParams(new String[]{"name"});
            rep.parse(context);
        }
        template.write(new FileOutputStream(new File("C:\\Users\\charles\\Desktop\\test.ppt")));
    }
}