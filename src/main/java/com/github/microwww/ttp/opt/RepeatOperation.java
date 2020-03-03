package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class RepeatOperation extends CopyOperation {

    @Override
    public void parse(ParseContext context) {
        super.parse(context, "copy");
    }

    @Override
    protected List<XSLFSlide> createSheet(ParseContext context, XSLFSlide sheet, List<Object> data) {
        List<XSLFSlide> shapes = new ArrayList<>();
        XMLSlideShow show = context.getTemplateShow();
        shapes.add(sheet);// 需要排序, 跟 data 的 setting 顺序一致
        for (int i = 1; i < data.size(); i++) {
            XSLFSlide slide = show.createSlide();
            slide = slide.importContent(sheet);
            shapes.add(slide);
        }
        return shapes;
    }

    @Override
    protected List<XSLFTextParagraph> createTextParagraphs(XSLFTextParagraph paragraph, List<Object> data) {
        int size = data.size();
        List<XSLFTextParagraph> res = new ArrayList<>(size);
        res.add(paragraph);
        for (int i = 1; i < size; i++) {
            res.add(Tools.copy(paragraph));
        }
        return res;
    }

    @Override
    protected List<XSLFTableRow> createTableRows(XSLFTable table, XSLFTableRow row, List<Object> data) {
        List<XSLFTableRow> shapes = new ArrayList<>();
        shapes.add(row);
        for (int i = 1; i < data.size(); i++) {
            shapes.add(Tools.copyTableRow(table, row));
        }
        return shapes;
    }

    @Override
    protected List<XSLFTable> createTables(XSLFSheet sheet, XSLFTable table, List<Object> data) {
        List<XSLFTable> shapes = new ArrayList<>();
        shapes.add(table);
        for (int i = 1; i < data.size(); i++) {
            XSLFTable target = Tools.copyTable(sheet, table);
            shapes.add(target);
        }
        return shapes;
    }
}
