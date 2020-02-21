package com.github.microwww.ttp.opt;

import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.*;

public class ReplaceOperationTest {

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

        ByteArrayOutputStream mem = new ByteArrayOutputStream();
        template.write(mem);

        try (InputStream in = new ByteArrayInputStream(mem.toByteArray())) {
            XMLSlideShow ppt = new XMLSlideShow(in);
            XSLFShape shape = ppt.getSlides().get(0).getShapes().get(0);
            XSLFTable table = (XSLFTable) shape;
            String txt = table.getRows().get(1).getCells().get(1).getText();
            assertEquals(map.get("name"), txt);
        }
        //try(FileOutputStream out = new FileOutputStream(new File("C:\\Users\\charles\\Desktop\\test.ppt"))) {
        //    out.write(mem.toByteArray());
        //}
    }
}