package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.util.DefaultMemberAccess;
import ognl.Ognl;
import ognl.OgnlContext;
import ognl.OgnlException;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.*;

import static org.junit.Assert.assertEquals;

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
        Map<String, Object> map = new HashMap<>();
        List<Map<String, String>> list = new ArrayList<>();
        list.add(Collections.singletonMap("name", "1"));
        list.add(Collections.singletonMap("name", "2"));
        list.add(Collections.singletonMap("name", "3"));
        list.add(Collections.singletonMap("name", "4"));
        list.add(Collections.singletonMap("name", "5"));
        list.add(Collections.singletonMap("name", "6"));
        map.put("list", list);
        Object value = Ognl.getValue("list.{name}", con, map);
        System.out.println("{  " + value + " }");
    }

    @Test
    public void copyTableRow() throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        XSLFSlide slide = template.getSlides().get(1);
        context.setTemplate(slide);

        XSLFTable table = null; // (XSLFTable) template.getSlides().get(1).getShapes().get(0);
        for (XSLFShape shapes : template.getSlides().get(1).getShapes()) {
            if (shapes instanceof XSLFTable) {
                table = (XSLFTable) shapes;
                break;
            }
        }

        int size = table.getRows().size();

        ArrayList<User> list = new ArrayList<>();
        list.add(new User("张三", 15));
        list.add(new User("李四", 11));
        list.add(new User("王五", 16));
        list.add(new User("赵六", 19));
        context.setData(Collections.singletonMap("list", list));

        RepeatOperation rep = new RepeatOperation();
        rep.setNode(new String[]{"XSLFTable", "0", "XSLFTableRow", "1"});
        rep.setParams(new String[]{"list", "null", "null", "null", "null", "item.name", "item.age", "index+1", "null", "null"});
        rep.parse(context);

        ByteArrayOutputStream mem = new ByteArrayOutputStream();
        template.write(mem);
        //try (FileOutputStream out = new FileOutputStream(new File("C:\\Users\\changshu.li\\Desktop\\test.pptx"))) {
        //    out.write(mem.toByteArray());
        //}
        try (InputStream in = new ByteArrayInputStream(mem.toByteArray())) {
            template = new XMLSlideShow(in);
        }

        for (XSLFShape shapes : template.getSlides().get(1).getShapes()) {
            if (shapes instanceof XSLFTable) {
                table = (XSLFTable) shapes;
                break;
            }
        }
        String text = table.getRows().get(size).getCells().get(4).getText();
        assertEquals(list.get(0).getName(), text);
        text = table.getRows().get(size).getCells().get(6).getText();
        assertEquals("1", text);
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

    public static class User {
        String name;
        int age;

        public User(String name, int age) {
            this.name = name;
            this.age = age;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }
    }
}