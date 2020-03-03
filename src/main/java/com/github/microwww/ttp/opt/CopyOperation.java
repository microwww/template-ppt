package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import com.github.microwww.ttp.util.DataUtil;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Stack;

public class CopyOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(DeleteOperation.class);

    @Override
    public void parse(ParseContext context) {
        super.parse(context, "copy");
    }

    public void copy(ParseContext context, Object o) {
        logger.warn("Not support copy type : {}", o.getClass());
    }

    public void copy(ParseContext context, XSLFSlide sheet) {
        // XSLFTextShape content = paragraph.getParentShape();
        ParamMessage[] param = this.getParamsWithPattern();
        Assert.isTrue(param.length > 0, "Repeat must has params");
        ParamMessage pm = param[0];
        List<Object> data = super.getCollectionValue(pm.getParam(), context.getDataStack());
        List<XSLFSlide> shapes = new ArrayList<>();
        XMLSlideShow show = context.getTemplateShow();
        shapes.add(sheet);// 需要排序, 跟 data 的 setting 顺序一致
        for (int i = 1; i < data.size(); i++) {
            XSLFSlide slide = show.createSlide();
            slide = slide.importContent(sheet);
            shapes.add(slide);
        }

        int index = sheet.getSlideNumber() - 1;
        Assert.isTrue(index >= 0, "slide index must >= 0");
        for (int k = 0; k < data.size(); k++) {
            int i = (k + 1) % data.size(); // 模板放到最后
            RepeatDomain info = new RepeatDomain();
            info.setItem(data.get(i));
            info.setIndex(i);
            next(context, shapes.get(i), info);
            show.setSlideOrder(shapes.get(i), index + i);
        }

    }

    public void copy(ParseContext context, XSLFTextParagraph paragraph) {
        // XSLFTextShape content = paragraph.getParentShape();
        ParamMessage[] param = this.getParamsWithPattern();
        Assert.isTrue(param.length > 0, "Repeat must has params");
        ParamMessage pm = param[0];
        Object os = super.getValue(pm.getParam(), context.getDataStack());
        List<Object> data = DataUtil.toList(os);
        int size = data.size();
        List<XSLFTextParagraph> res = new ArrayList<>(size);
        res.add(paragraph);
        for (int i = 1; i < size; i++) {
            res.add(Tools.copy(paragraph));
        }
        for (int i = 1; i <= size; i++) {
            int k = i % size;
            Tools.replace(res.get(k), paragraph.getText(), pm.format(data.get(k)));
        }
    }

    public void copy(ParseContext context, XSLFTableRow row) {
        Stack<Object> stack = context.getContainer();
        XSLFTable table = (XSLFTable) stack.get(stack.size() - 2);
        // XSLFTable table = (XSLFTable) context.getContainer().peek();
        //XSLFTable table = DeleteOperation.getTable(context.getTemplate(), row);
        String[] param = this.getParams();
        Assert.isTrue(param.length > 0, "repeat XSLFTable must have [count]");
        List<Object> data = super.getCollectionValue(param[0], context.getDataStack());
        //ItemInfo item = new ItemInfo();
        List<XSLFTableRow> shapes = new ArrayList<>();
        shapes.add(row);
        for (int i = 1; i < data.size(); i++) {
            shapes.add(Tools.copyTableRow(table, row));
        }

        for (int k = 0; k < data.size(); k++) {
            int i = (k + 1) % data.size(); // 模板放到最后
            RepeatDomain info = new RepeatDomain();
            info.setItem(data.get(i));
            info.setIndex(i);
            next(context, shapes.get(i), info);
        }
        if (data.isEmpty()) {
            int size = table.getRows().size();
            for (int i = 0; i < size; i++) {
                XSLFTableRow tr = table.getRows().get(i);
                if (tr.equals(row)) {
                    table.removeRow(i);
                    break;
                }
            }
        }
    }

    public void copy(ParseContext context, XSLFTable table) {
        XSLFSheet sheet = context.getTemplate();
        String[] param = this.getParams();
        Assert.isTrue(param.length > 1, "repeat XSLFTable must have [list/array, position], tow param");
        List<Object> data;
        try {
            data = super.getCollectionValue(param[0], context.getDataStack());
        } catch (RuntimeException e) {// ignore
            throw new RuntimeException("FIRST param is list/array !", e);
        }
        String[] ps = param[1].split(",");
        Assert.isTrue(ps.length == 2, "Repeat position.split(',') != 2");
        List<XSLFTable> shapes = new ArrayList<>();
        shapes.add(table);
        for (int i = 1; i < data.size(); i++) {
            XSLFTable target = Tools.copyTable(sheet, table);
            Rectangle2D anchor = target.getAnchor();
            //anchor = _Help.rectanglePx2point(anchor, 0,0,0,0);
            double x = Double.valueOf(ps[0]).doubleValue() * (i);
            double y = Double.valueOf(ps[1]).doubleValue() * (i);
            Rectangle2D.Double r2d = new Rectangle2D.Double(anchor.getX() + x, anchor.getY() + y, anchor.getWidth() + x, anchor.getHeight() + y);
            //anchor.add(0, i * 2);
            target.setAnchor(r2d);
            shapes.add(target);
        }
        for (int k = 0; k < data.size(); k++) {
            int i = (k + 1) % data.size(); // 模板放到最后
            RepeatDomain info = new RepeatDomain();
            info.setItem(data.get(i));
            info.setIndex(i);
            next(context, shapes.get(i), info);
        }
    }

    private void next(ParseContext context, Object shape, RepeatDomain info) {
        try {
            context.getContainer().push(shape);
            context.getDataStack().push(info);
            for (Operation childrenOperation : this.childrenOperations) {
                childrenOperation.setParentOperations(this);
                childrenOperation.parse(context);
            }
        } finally {
            context.getDataStack().pop();
            context.getContainer().pop();
        }
    }
}
