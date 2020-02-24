package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.geom.Rectangle2D;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Stack;

public class RepeatOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(DeleteOperation.class);

    private static final ThreadLocal<Stack> stack = new ThreadLocal() {
        @Override
        protected Stack initialValue() {
            return new Stack<>();
        }
    };

    public static Stack repeatStark() {
        return stack.get();
    }

    @Override
    public void parse(ParseContext context) {
        List<?> search = super.search(context);
        for (Object item : search) {
            thisInvoke("copy", context, item);
        }
    }

    public void copy(ParseContext context, Object o) {
        logger.warn("Not support copy type : {}", o.getClass());
    }

    public void copy(ParseContext context, XSLFTextParagraph paragraph) {
        // XSLFTextShape content = paragraph.getParentShape();
        ParamMessage[] param = this.getParamsWithPattern();
        Assert.isTrue(param.length > 0, "Repeat must has params");
        ParamMessage pm = param[0];
        Object os = super.getValue(pm.getParam(), context.getData());
        Collection<Object> list = ReplaceOperation.toList(os);
        for (Object val : list) {
            XSLFTextParagraph copy = Tools.copy(paragraph);
            Tools.replace(copy, paragraph.getText(), pm.format(val));
        }
    }

    public void copy(ParseContext context, XSLFTableRow row) {
        XSLFTable table = DeleteOperation.getTable(context.getTemplate(), row);
        String[] param = this.getParams();
        Assert.isTrue(param.length > 0, "repeat XSLFTable must have [count]");
        List<Object> list = super.getValue(param[0], context.getData(), List.class);
        for (int i = 0; i < list.size(); i++) {
            XSLFTableRow nrow = Tools.copyTableRow(table, row);
            List<XSLFTableCell> cells = nrow.getCells();
            for (int k = 0; k < cells.size(); k++) {
                if (k + 1 < param.length) {
                    String exp = param[k + 1];
                    XSLFTableCell cell = cells.get(k);
                    repeat(context, list.get(i), exp, cell);
                }
            }

        }
    }

    public void repeat(ParseContext context, Object obj, String exp, XSLFTextShape cell) {
        ReplaceOperation rp = new ReplaceOperation();
        if (exp.equalsIgnoreCase("null")) {
            rp.setParams(new String[]{});
        } else {
            rp.setParams(new String[]{exp});
        }
        Object origin = context.getData();
        try {
            context.setData(Collections.singletonMap("item", obj));
            rp.replace(context, cell);
        } catch (Exception e) {// ignore
            context.setData(origin);
            rp.replace(context, cell);
        } finally {
            context.setData(origin);
        }
    }

    public void copy(ParseContext context, XSLFTable table) {
        XSLFSheet sheet = context.getTemplate();
        String[] param = this.getParams();
        Assert.isTrue(param.length > 1, "repeat XSLFTable must have [count, position], tow param");
        int count = Integer.valueOf(super.getValue(param[0], context.getData(), String.class)).intValue();
        String[] ps = param[1].split(",");
        Assert.isTrue(ps.length == 2, "Repeat position.split(',') != 2");
        for (int i = 0; i < count; i++) {
            XSLFTable target = Tools.copyTable(sheet, table);
            Rectangle2D anchor = target.getAnchor();
            //anchor = _Help.rectanglePx2point(anchor, 0,0,0,0);
            double x = Double.valueOf(ps[0]).doubleValue() * (i + 1);
            double y = Double.valueOf(ps[1]).doubleValue() * (i + 1);
            Rectangle2D.Double r2d = new Rectangle2D.Double(anchor.getX() + x, anchor.getY() + y, anchor.getWidth() + x, anchor.getHeight() + y);
            //anchor.add(0, i * 2);
            target.setAnchor(r2d);
        }
    }
}
