package com.github.microwww.ttp.replace;

import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import static org.junit.Assert.*;

public class SearchTableCellTest {
    public static final String PATH = _HelpTest.PATH;

    @Test
    public void testSearch() throws IOException {
        XMLSlideShow target, template;
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            template = new XMLSlideShow(in);
        }
        try (FileInputStream in = new FileInputStream(new File(PATH, "template.pptx"))) {
            target = new XMLSlideShow(in);
            for (int i = target.getSlides().size(); i > 0; i--) {
                target.removeSlide(i - 1);
            }
        }
        XSLFSlide slide = target.createSlide();
        XSLFTable table = Tools.copyTable(slide, (XSLFTable) template.getSlides().get(0).getShapes().get(0));
        List<TextExpress> search = new SearchTable(table).search();
        Map<String, Object> map = new HashMap<>();
        map.put("name", "china");
        map.put("age", 11);
        for(TextExpress run: search){
            Object val = map.get(run.getExpress());
            run.replace(val.toString());
        }

        File file = new File(PATH, UUID.randomUUID().toString() + ".pptx");
        target.write(new FileOutputStream(file));

        try (FileInputStream in = new FileInputStream(file)) {
            XSLFShape shape = new XMLSlideShow(in).getSlides().get(0).getShapes().get(0);
            XSLFTable tb = (XSLFTable) shape;
            String text = tb.getRows().get(3).getCells().get(2).getText();
            assertTrue(text.contains("china"));
            assertTrue(text.contains("11"));
        }
    }
}