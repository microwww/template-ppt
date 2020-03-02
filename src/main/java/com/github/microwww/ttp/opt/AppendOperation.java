package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Tools;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Stack;

public class AppendOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(AppendOperation.class);

    @Override
    public void parse(ParseContext context) {
        super.parse(context, "append");
    }

    public void append(Object o) {
        logger.warn("Not support delete type : {}", o.getClass());
    }

    public void append(ParseContext context, XSLFTableRow row) {
        Stack<Object> stack = context.getContainer();
        XSLFTable table = (XSLFTable) stack.get(stack.size() - 2);
        XSLFTableRow newRow = Tools.copyTableRow(table, row);
        List<XSLFTableCell> cells = newRow.getCells();
        ParamMessage[] pms = this.getParamsWithPattern();
        for (int i = 0; i < cells.size() && i < pms.length; i++) {
            StringBuilder buffer = new StringBuilder();
            ParamMessage param = pms[i];
            Object val = getValue(param.getParam(), context.getDataStack());
            buffer.append(param.format(val));
            Tools.setTextShapeWithStyle(cells.get(i), buffer.toString());
        }
    }
}
