package com.github.microwww.ttp.opt;

import com.github.microwww.ttp.Assert;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class MergeOperation extends Operation {
    private static final Logger logger = LoggerFactory.getLogger(MergeOperation.class);

    @Override
    public void parse(ParseContext context) {
        List<?> nodes = this.search(context);
        for (Object node : nodes) {
            thisInvoke("merge", context, node);
        }
    }

    public void merge(Object o) {
        logger.warn("Not support delete type : {}", o.getClass());
    }


    public void merge(ParseContext context, XSLFTable table) {
        String[] params = this.getParams();
        Assert.isTrue(params.length >= 1, "Please set col");
        int col = this.getValue(params[0], context.getDataStack(), Integer.class).intValue();
        if (!table.getRows().isEmpty()) {
            List<XSLFTableCell> cells = table.getRows().get(0).getCells();
            if (!cells.isEmpty() && cells.size() > col) {
                String text = "";
                List<XSLFTableRow> rows = table.getRows();
                int from = 0;
                for (int i = from; i < rows.size(); i++) {
                    String trim = rows.get(i).getCells().get(col).getText().trim();
                    if (!trim.equals(text)) {
                        if (i - from > 1) {
                            table.mergeCells(from, i - 1, col, col);
                        }
                        text = trim;
                        from = i;
                    } else {
                        if (i == rows.size() - 1 && i - from > 1) {
                            table.mergeCells(from, i, col, col);
                        }
                    }
                }

            }
        }
    }
}
