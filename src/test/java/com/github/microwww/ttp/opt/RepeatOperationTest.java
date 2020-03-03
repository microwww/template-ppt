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
        ParseContext context = createContext(0);
        XMLSlideShow template = context.getTemplateShow();
        XSLFTable table = getTable(context, 1, 0);
        List<User> list = getDemoList();
        int size = table.getRows().size();
        context.setData(Collections.singletonMap("list", list));

        CopyOperation rep = new CopyOperation();
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
        assertEquals(list.get(0).name, text);
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

        CopyOperation rep = new CopyOperation();
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

    @Test
    public void repeatTable() throws IOException {
        ParseContext context = createContext(1);
        XMLSlideShow template = context.getTemplateShow();
        //XSLFTable table = getTable(context, 1, 0);
        context.addData("list", Arrays.asList(new Group("亚洲", getDemoList()), new Group("欧洲", getDemoList()), new Group("非洲", getDemoList())));

        CopyOperation rep = new CopyOperation();
        rep.setNode(new String[]{"XSLFTable", "0"});
        rep.setParams(new String[]{"list", "0,100"});
        {
            ReplaceOperation rep2 = new ReplaceOperation();
            rep2.setNode(new String[]{"XSLFTableRow", "0", "XSLFTableCell", "0"});
            rep2.setParams(new String[]{"item.name"});
            rep.addChildrenOperation(rep2);
        }
        {
            CopyOperation rep2 = new CopyOperation();
            rep2.setNode(new String[]{"XSLFTableRow", "1"});
            rep2.setParams(new String[]{"item.users"});
            rep.addChildrenOperation(rep2);
            {
                ReplaceOperation rep3 = new ReplaceOperation();
                rep3.setNode(new String[]{"XSLFTableCell", "6"});
                rep3.setParams(new String[]{"item.name"});
                rep2.addChildrenOperation(rep3);
            }
            {
                ReplaceOperation rep3 = new ReplaceOperation();
                rep3.setNode(new String[]{"XSLFTableCell", "5"});
                rep3.setParams(new String[]{"item.age"});
                rep2.addChildrenOperation(rep3);
            }
        }
        {
            DeleteOperation rep2 = new DeleteOperation();
            rep2.setNode(new String[]{"XSLFTableRow", "1-4"});
            rep2.setParams(new String[]{});
            rep.addChildrenOperation(rep2);
        }
        rep.parse(context);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        template.write(bos);
        {
            template = new XMLSlideShow(new ByteArrayInputStream(bos.toByteArray()));
            template.getSlides().get(1);
        }

        try (FileOutputStream out = new FileOutputStream(new File("C:\\Users\\charles\\Desktop\\test.pptx"))) {
            template.write(out);
        }
    }

    public static class Group {
        public final String name;
        public final List<User> users;

        public Group(String name, List<User> users) {
            this.name = name;
            this.users = users;
        }

    }

    public static class User {
        public final String name;
        public final int age;

        public User(String name, int age) {
            this.name = name;
            this.age = age;
        }
    }

    public ParseContext createContext(int slide) throws IOException {
        XMLSlideShow template;
        try (FileInputStream in = new FileInputStream(new File(_HelpTest.PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        ParseContext context = new ParseContext(template);
        context.setTemplate(template.getSlides().get(slide));
        return context;
    }

    public XSLFTable getTable(ParseContext context, int slide, int index) {
        XMLSlideShow template = context.getTemplateShow();
        XSLFTable table = null; // (XSLFTable) template.getSlides().get(1).getShapes().get(0);
        int idx = 0;
        for (XSLFShape shapes : template.getSlides().get(slide).getShapes()) {
            if (shapes instanceof XSLFTable) {
                if (index == idx) {
                    return (XSLFTable) shapes;
                }
            }
        }
        throw new RuntimeException("Not find !");
    }

    public List<User> getDemoList() {
        ArrayList<User> list = new ArrayList<>();
        list.add(new User("张三", (int) (Math.random() * 100)));
        list.add(new User("李四", (int) (Math.random() * 100)));
        list.add(new User("王五", (int) (Math.random() * 100)));
        list.add(new User("赵六", (int) (Math.random() * 100)));
        return list;
    }
}