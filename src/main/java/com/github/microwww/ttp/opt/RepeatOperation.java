package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.geom.Rectangle2D;
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
        Assert.isTrue(nodeStack.size() >= 2, "STACK MUST DEPEND > 2");
        List<Object> peek = nodeStack.get(nodeStack.size() - 2);
        XSLFTextShape content = null;
        int i = 0;
        for (Object pk : peek) {
            if (pk instanceof XSLFTextShape) {
                XSLFTextShape shape = (XSLFTextShape) pk;
                i = 0;
                for (XSLFTextParagraph gh : shape.getTextParagraphs()) {
                    i++;
                    if (gh.equals(paragraph)) {
                        content = shape;
                        break;
                    }
                }
            }
            if (content != null) {
                break;
            }
        }
        if (content != null) {
            ParamMessage[] param = this.getParamsWithPattern();
            for (int k = 0; k < param.length; k++) {
                ParamMessage pm = param[k];
                Object val = super.getValue(pm.getParam(), context.getData());
                XSLFTextParagraph tp = content.addNewTextParagraph();
                tp.getXmlObject().set(paragraph.getXmlObject().copy());
                content.getTextBody().getParagraph(i + k).setText(pm.format(val));
            }
        }
    }

    public void copy(ParseContext context, XSLFTableRow row) {
        XSLFTable table = DeleteOperation.getTable(context.getTemplate(), row);
        String[] param = this.getParams();
        Assert.isTrue(param.length > 0, "repeat XSLFTable must have [count]");
        List<Object> list = super.getCollectionValue(param[0], context.getData());
        ItemInfo item = new ItemInfo(context);
        for (int i = 0; i < list.size(); i++) {
            XSLFTableRow nrow = Tools.copyTableRow(table, row);
            List<XSLFTableCell> cells = nrow.getCells();
            for (int k = 0; k < cells.size(); k++) {
                if (k + 1 < param.length) {
                    String exp = param[k + 1];
                    XSLFTableCell cell = cells.get(k);
                    item.setIndex(i).setItem(list.get(i));
                    repeat(item, exp, cell);
                }
            }

        }
    }

    public void repeat(ItemInfo item, String exp, XSLFTextShape cell) {
        ReplaceOperation rp = new ReplaceOperation();
        if (exp.equalsIgnoreCase("null")) {
            rp.setParams(new String[]{});
        } else {
            rp.setParams(new String[]{exp});
        }
        Object origin = item.context.getData();
        try {
            item.context.setData(item);
            rp.replace(item.context, cell);
        } catch (Exception e) {// ignore
            item.context.setData(origin);
            rp.replace(item.context, cell);
        } finally {
            item.context.setData(origin);
        }
    }

    public static class ItemInfo {
        private final ParseContext context;

        private Object item;
        private int index;

        public ItemInfo(ParseContext context) {
            this.context = context;
        }

        public Object getItem() {
            return item;
        }

        public ItemInfo setItem(Object item) {
            this.item = item;
            return this;
        }

        public int getIndex() {
            return index;
        }

        public ItemInfo setIndex(int index) {
            this.index = index;
            return this;
        }
    }

    public void copy(ParseContext context, XSLFTable table) {
        XSLFSheet sheet = context.getTemplate();
        String[] param = this.getParams();
        Assert.isTrue(param.length > 1, "repeat XSLFTable must have [count, position], tow param");
        int count = super.getValue(param[0], context.getData(), Integer.class).intValue();
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
