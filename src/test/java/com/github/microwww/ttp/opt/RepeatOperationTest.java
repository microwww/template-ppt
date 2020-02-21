package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.util.DefaultMemberAccess;
import ognl.Ognl;
import ognl.OgnlContext;
import ognl.OgnlException;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.List;

import static org.junit.Assert.*;

public class RepeatOperationTest {

    @Test
    public void getElement() {
    }

    @Test
    public void searchRanges() {
        {
            List<Range> ranges = Operation.searchRanges("1,5");
            assertEquals(ranges.size(), 2);
            assertEquals(ranges.get(0).getFrom(), 1);
            assertEquals(ranges.get(0).getTo(), 2);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 6);
        }
        {
            List<Range> ranges = Operation.searchRanges("1-");
            assertEquals(ranges.size(), 1);
            assertEquals(ranges.get(0).getFrom(), 1);
            assertEquals(ranges.get(0).getTo(), Integer.MAX_VALUE);
        }
        {
            List<Range> ranges = Operation.searchRanges("2,5-8,11");
            assertEquals(ranges.size(), 3);
            assertEquals(ranges.get(0).getFrom(), 2);
            assertEquals(ranges.get(0).getTo(), 3);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 8);
            assertEquals(ranges.get(2).getFrom(), 11);
            assertEquals(ranges.get(2).getTo(), 12);
        }
        {
            List<Range> ranges = Operation.searchRanges("2,5-8,11-20,30-");
            assertEquals(ranges.size(), 4);
            assertEquals(ranges.get(0).getFrom(), 2);
            assertEquals(ranges.get(0).getTo(), 3);
            assertEquals(ranges.get(1).getFrom(), 5);
            assertEquals(ranges.get(1).getTo(), 8);
            assertEquals(ranges.get(2).getFrom(), 11);
            assertEquals(ranges.get(2).getTo(), 20);
            assertEquals(ranges.get(3).getFrom(), 30);
            assertEquals(ranges.get(3).getTo(), Integer.MAX_VALUE);
        }

    }

    @Test
    public void ctest() throws IOException, OgnlException {
        OgnlContext con = new OgnlContext(null, null, new DefaultMemberAccess(true));
        Object value = Ognl.getValue("5", con, new HashMap<>());
        System.out.println("{  " + value + " }");
    }

    @Test
    public void copyTable() throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        XSLFSlide slide = template.getSlides().get(0);
        context.setTemplate(slide);

        RepeatOperation rep = new RepeatOperation();
        rep.setNode(new String[]{"XSLFTable", "0"});
        rep.setParams(new String[]{"'2'", "0,100"});
        rep.parse(context);

        ByteArrayOutputStream mem = new ByteArrayOutputStream();
        template.write(mem);
        try (FileOutputStream out = new FileOutputStream(new File("C:\\Users\\charles\\Desktop\\test.ppt"))) {
            out.write(mem.toByteArray());
        }
    }

    @Test
    public void copyRow() throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        XSLFSlide slide = template.getSlides().get(0);
        context.setTemplate(slide);
        int size = ((XSLFTable) slide.getShapes().get(0)).getRows().size();

        RepeatOperation rep = new RepeatOperation();
        rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "1"});
        rep.setParams(new String[]{"'2'"});
        rep.parse(context);

        ByteArrayOutputStream mem = new ByteArrayOutputStream();
        template.write(mem);

        try (InputStream in = new ByteArrayInputStream(mem.toByteArray())) {
            XMLSlideShow ppt = new XMLSlideShow(in);
            XSLFShape shape = ppt.getSlides().get(0).getShapes().get(0);
            XSLFTable table = (XSLFTable) shape;
            assertEquals(size + 2, table.getRows().size());
        }
    }
}